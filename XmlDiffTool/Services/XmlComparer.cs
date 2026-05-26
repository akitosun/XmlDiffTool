using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using XmlDiffTool.Models;

namespace XmlDiffTool.Services
{
    public class XmlComparer
    {
        public XmlComparisonResult Compare(string leftPath, string rightPath, bool ignoreCase, IProgress<int>? progress = null)
        {
            progress?.Report(0);

            var leftDocument = LoadDocument(leftPath);
            progress?.Report(35);

            var rightDocument = LoadDocument(rightPath);
            progress?.Report(70);

            var options = new CompareOptions(ignoreCase);
            var roots = CompareElements(leftDocument.Root, rightDocument.Root, string.Empty, options)
                .ToList();

            progress?.Report(100);
            return new XmlComparisonResult(roots);
        }

        private static XDocument LoadDocument(string path)
        {
            using var stream = File.OpenRead(path);
            return XDocument.Load(stream, LoadOptions.SetLineInfo);
        }

        private static IEnumerable<XmlDifferenceNode> CompareElements(XElement? left, XElement? right, string parentPath, CompareOptions options)
        {
            if (left is null && right is null)
            {
                yield break;
            }

            if (left is null)
            {
                yield return CreateMissingElementNode(right!, parentPath, isLeftMissing: true);
                yield break;
            }

            if (right is null)
            {
                yield return CreateMissingElementNode(left, parentPath, isRightMissing: true);
                yield break;
            }

            var leftName = left.Name.LocalName;
            var rightName = right.Name.LocalName;
            if (!options.NamesEqual(leftName, rightName))
            {
                yield return CreateMissingElementNode(left, parentPath, isRightMissing: true);
                yield return CreateMissingElementNode(right, parentPath, isLeftMissing: true);
                yield break;
            }

            var path = AppendPath(parentPath, leftName);
            var node = new XmlDifferenceNode(path, leftName, XmlDifferenceKind.Element);

            AddAttributeDifferences(node, left, right, path, options);
            AddValueDifference(node, left, right, path, options);
            AddChildDifferences(node, left, right, path, options);

            if (node.HasChildren)
            {
                yield return node;
            }
        }

        private static void AddAttributeDifferences(XmlDifferenceNode node, XElement left, XElement right, string path, CompareOptions options)
        {
            var leftAttributes = left.Attributes()
                .Where(attribute => !attribute.IsNamespaceDeclaration)
                .GroupBy(attribute => attribute.Name.LocalName, options.NameComparer)
                .ToDictionary(group => group.Key, group => group.First().Value, options.NameComparer);
            var rightAttributes = right.Attributes()
                .Where(attribute => !attribute.IsNamespaceDeclaration)
                .GroupBy(attribute => attribute.Name.LocalName, options.NameComparer)
                .ToDictionary(group => group.Key, group => group.First().Value, options.NameComparer);

            var names = new HashSet<string>(leftAttributes.Keys, options.NameComparer);
            names.UnionWith(rightAttributes.Keys);

            foreach (var name in names.OrderBy(name => name, StringComparer.OrdinalIgnoreCase))
            {
                leftAttributes.TryGetValue(name, out var leftValue);
                rightAttributes.TryGetValue(name, out var rightValue);

                var isLeftMissing = !leftAttributes.ContainsKey(name);
                var isRightMissing = !rightAttributes.ContainsKey(name);
                if (!isLeftMissing && !isRightMissing && options.ValuesEqual(leftValue, rightValue))
                {
                    continue;
                }

                node.Children.Add(new XmlDifferenceNode(
                    $"{path}[@{name}]",
                    $"@{name}",
                    XmlDifferenceKind.Attribute,
                    leftValue,
                    rightValue,
                    isLeftMissing,
                    isRightMissing,
                    isLeftMissing ? null : GetLineNumber(left.Attributes().FirstOrDefault(attribute => !attribute.IsNamespaceDeclaration && options.NamesEqual(attribute.Name.LocalName, name))),
                    isRightMissing ? null : GetLineNumber(right.Attributes().FirstOrDefault(attribute => !attribute.IsNamespaceDeclaration && options.NamesEqual(attribute.Name.LocalName, name)))));
            }
        }

        private static void AddValueDifference(XmlDifferenceNode node, XElement left, XElement right, string path, CompareOptions options)
        {
            if (GetComparableChildren(left).Any() || GetComparableChildren(right).Any())
            {
                return;
            }

            var leftValue = NormalizeText(left.Value);
            var rightValue = NormalizeText(right.Value);
            if (options.ValuesEqual(leftValue, rightValue))
            {
                return;
            }

            node.Children.Add(new XmlDifferenceNode(
                $"{path}[#value]",
                "#value",
                XmlDifferenceKind.Value,
                leftValue,
                rightValue,
                leftLineNumber: GetLineNumber(left),
                rightLineNumber: GetLineNumber(right)));
        }

        private static void AddChildDifferences(XmlDifferenceNode node, XElement left, XElement right, string path, CompareOptions options)
        {
            var leftChildren = GetComparableChildren(left);
            var rightChildren = GetComparableChildren(right);
            if (leftChildren.Count == 0 && rightChildren.Count == 0)
            {
                return;
            }

            var childNames = new HashSet<string>(leftChildren.Select(child => child.Name.LocalName), options.NameComparer);
            childNames.UnionWith(rightChildren.Select(child => child.Name.LocalName));

            foreach (var childName in childNames.OrderBy(name => name, StringComparer.OrdinalIgnoreCase))
            {
                var leftGroup = leftChildren.Where(child => options.NamesEqual(child.Name.LocalName, childName)).ToList();
                var rightGroup = rightChildren.Where(child => options.NamesEqual(child.Name.LocalName, childName)).ToList();
                RemoveExactMatches(leftGroup, rightGroup, options);

                leftGroup.Sort((first, second) => string.Compare(BuildSignature(first, options), BuildSignature(second, options), StringComparison.Ordinal));
                rightGroup.Sort((first, second) => string.Compare(BuildSignature(first, options), BuildSignature(second, options), StringComparison.Ordinal));

                var pairedCount = Math.Min(leftGroup.Count, rightGroup.Count);
                for (var index = 0; index < pairedCount; index++)
                {
                    node.Children.AddRange(CompareElements(leftGroup[index], rightGroup[index], path, options));
                }

                for (var index = pairedCount; index < leftGroup.Count; index++)
                {
                    node.Children.Add(CreateMissingElementNode(leftGroup[index], path, isRightMissing: true));
                }

                for (var index = pairedCount; index < rightGroup.Count; index++)
                {
                    node.Children.Add(CreateMissingElementNode(rightGroup[index], path, isLeftMissing: true));
                }
            }
        }

        private static void RemoveExactMatches(List<XElement> leftGroup, List<XElement> rightGroup, CompareOptions options)
        {
            var rightBySignature = rightGroup
                .GroupBy(child => BuildSignature(child, options), StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => new Queue<XElement>(group), StringComparer.Ordinal);

            for (var index = leftGroup.Count - 1; index >= 0; index--)
            {
                var signature = BuildSignature(leftGroup[index], options);
                if (!rightBySignature.TryGetValue(signature, out var matches) || matches.Count == 0)
                {
                    continue;
                }

                var matchedRight = matches.Dequeue();
                rightGroup.Remove(matchedRight);
                leftGroup.RemoveAt(index);
            }
        }

        private static XmlDifferenceNode CreateMissingElementNode(XElement element, string parentPath, bool isLeftMissing = false, bool isRightMissing = false)
        {
            var path = AppendPath(parentPath, element.Name.LocalName);
            var node = new XmlDifferenceNode(
                path,
                element.Name.LocalName,
                XmlDifferenceKind.Element,
                isLeftMissing ? null : NormalizeText(element.Value),
                isRightMissing ? null : NormalizeText(element.Value),
                isLeftMissing,
                isRightMissing,
                isLeftMissing ? null : GetLineNumber(element),
                isRightMissing ? null : GetLineNumber(element));

            foreach (var attribute in element.Attributes()
                         .Where(attribute => !attribute.IsNamespaceDeclaration)
                         .OrderBy(attribute => attribute.Name.LocalName, StringComparer.OrdinalIgnoreCase))
            {
                node.Children.Add(new XmlDifferenceNode(
                    $"{path}[@{attribute.Name.LocalName}]",
                    $"@{attribute.Name.LocalName}",
                    XmlDifferenceKind.Attribute,
                    isLeftMissing ? null : attribute.Value,
                    isRightMissing ? null : attribute.Value,
                    isLeftMissing,
                    isRightMissing,
                    isLeftMissing ? null : GetLineNumber(attribute),
                    isRightMissing ? null : GetLineNumber(attribute)));
            }

            var childElements = GetComparableChildren(element);
            if (childElements.Count == 0)
            {
                var value = NormalizeText(element.Value);
                if (!string.IsNullOrEmpty(value))
                {
                    node.Children.Add(new XmlDifferenceNode(
                        $"{path}[#value]",
                        "#value",
                        XmlDifferenceKind.Value,
                        isLeftMissing ? null : value,
                        isRightMissing ? null : value,
                        isLeftMissing,
                        isRightMissing,
                        isLeftMissing ? null : GetLineNumber(element),
                        isRightMissing ? null : GetLineNumber(element)));
                }

                return node;
            }

            node = new XmlDifferenceNode(
                path,
                element.Name.LocalName,
                XmlDifferenceKind.Element,
                isLeftMissing: isLeftMissing,
                isRightMissing: isRightMissing,
                leftLineNumber: isLeftMissing ? null : GetLineNumber(element),
                rightLineNumber: isRightMissing ? null : GetLineNumber(element));

            foreach (var attribute in element.Attributes()
                         .Where(attribute => !attribute.IsNamespaceDeclaration)
                         .OrderBy(attribute => attribute.Name.LocalName, StringComparer.OrdinalIgnoreCase))
            {
                node.Children.Add(new XmlDifferenceNode(
                    $"{path}[@{attribute.Name.LocalName}]",
                    $"@{attribute.Name.LocalName}",
                    XmlDifferenceKind.Attribute,
                    isLeftMissing ? null : attribute.Value,
                    isRightMissing ? null : attribute.Value,
                    isLeftMissing,
                    isRightMissing,
                    isLeftMissing ? null : GetLineNumber(attribute),
                    isRightMissing ? null : GetLineNumber(attribute)));
            }

            foreach (var child in childElements.OrderBy(child => child.Name.LocalName, StringComparer.OrdinalIgnoreCase))
            {
                node.Children.Add(CreateMissingElementNode(child, path, isLeftMissing, isRightMissing));
            }

            return node;
        }

        private static string BuildSignature(XElement element, CompareOptions options)
        {
            var builder = new StringBuilder();
            builder.Append(options.NormalizeName(element.Name.LocalName));

            foreach (var attribute in element.Attributes()
                         .Where(attribute => !attribute.IsNamespaceDeclaration)
                         .OrderBy(attribute => options.NormalizeName(attribute.Name.LocalName), StringComparer.Ordinal))
            {
                builder.Append("|@");
                builder.Append(options.NormalizeName(attribute.Name.LocalName));
                builder.Append('=');
                builder.Append(options.NormalizeValue(attribute.Value));
            }

            var childElements = GetComparableChildren(element);
            if (childElements.Count == 0)
            {
                builder.Append("|#=");
                builder.Append(options.NormalizeValue(NormalizeText(element.Value)));
            }

            foreach (var childSignature in childElements
                         .Select(child => BuildSignature(child, options))
                         .OrderBy(signature => signature, StringComparer.Ordinal))
            {
                builder.Append("|<");
                builder.Append(childSignature);
                builder.Append('>');
            }

            return builder.ToString();
        }

        private static List<XElement> GetComparableChildren(XElement element)
        {
            var children = element.Elements().ToList();
            if (children.Count > 0)
            {
                return children;
            }

            return TryParseEmbeddedXml(NormalizeText(element.Value), GetLineNumber(element), out var embeddedChildren)
                ? embeddedChildren
                : children;
        }

        private static bool TryParseEmbeddedXml(string value, int? parentLineNumber, out List<XElement> children)
        {
            children = new List<XElement>();
            if (string.IsNullOrWhiteSpace(value) || !value.TrimStart().StartsWith("<", StringComparison.Ordinal))
            {
                return false;
            }

            if (TryParseDocument(value, out var document) && document.Root is not null)
            {
                children.Add(document.Root);
                AnnotateEmbeddedLineNumbers(children, parentLineNumber);
                return true;
            }

            if (TryParseDocument($"<__xml_diff_embedded_root>{value}</__xml_diff_embedded_root>", out document)
                && document.Root is not null)
            {
                children.AddRange(document.Root.Elements());
                AnnotateEmbeddedLineNumbers(children, parentLineNumber);
                return children.Count > 0;
            }

            return false;
        }

        private static bool TryParseDocument(string value, out XDocument document)
        {
            try
            {
                document = XDocument.Parse(value, LoadOptions.SetLineInfo);
                return true;
            }
            catch (XmlException)
            {
                document = new XDocument();
                return false;
            }
        }

        private static void AnnotateEmbeddedLineNumbers(IEnumerable<XElement> elements, int? parentLineNumber)
        {
            foreach (var element in elements)
            {
                AnnotateEmbeddedLineNumber(element, parentLineNumber);
            }
        }

        private static void AnnotateEmbeddedLineNumber(XElement element, int? parentLineNumber)
        {
            AnnotateEmbeddedLineNumber((XObject)element, parentLineNumber);

            foreach (var attribute in element.Attributes())
            {
                AnnotateEmbeddedLineNumber(attribute, parentLineNumber);
            }

            foreach (var child in element.Elements())
            {
                AnnotateEmbeddedLineNumber(child, parentLineNumber);
            }
        }

        private static void AnnotateEmbeddedLineNumber(XObject item, int? parentLineNumber)
        {
            var lineNumber = GetNativeLineNumber(item);
            if (parentLineNumber is not null && lineNumber is not null)
            {
                item.AddAnnotation(new SourceLineNumber(parentLineNumber.Value + lineNumber.Value - 1));
            }
        }

        private static int? GetLineNumber(XObject? item)
        {
            return item?.Annotation<SourceLineNumber>()?.LineNumber ?? GetNativeLineNumber(item);
        }

        private static int? GetNativeLineNumber(XObject? item)
        {
            if (item is IXmlLineInfo lineInfo && lineInfo.HasLineInfo())
            {
                return lineInfo.LineNumber;
            }

            return null;
        }

        private static string AppendPath(string parentPath, string name)
        {
            return string.IsNullOrWhiteSpace(parentPath) ? name : $"{parentPath}/{name}";
        }

        private static string NormalizeText(string value)
        {
            return value.Trim();
        }

        private sealed class CompareOptions
        {
            public CompareOptions(bool ignoreCase)
            {
                IgnoreCase = ignoreCase;
                NameComparer = ignoreCase ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;
            }

            public bool IgnoreCase { get; }

            public StringComparer NameComparer { get; }

            public bool NamesEqual(string left, string right)
            {
                return string.Equals(left, right, IgnoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);
            }

            public bool ValuesEqual(string? left, string? right)
            {
                return string.Equals(left, right, IgnoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);
            }

            public string NormalizeName(string value)
            {
                return IgnoreCase ? value.ToUpperInvariant() : value;
            }

            public string NormalizeValue(string value)
            {
                return IgnoreCase ? value.ToUpperInvariant() : value;
            }
        }

        private sealed class SourceLineNumber
        {
            public SourceLineNumber(int lineNumber)
            {
                LineNumber = lineNumber;
            }

            public int LineNumber { get; }
        }
    }
}
