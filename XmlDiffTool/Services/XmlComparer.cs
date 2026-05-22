using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
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
            return XDocument.Load(stream, LoadOptions.None);
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
                    isRightMissing));
            }
        }

        private static void AddValueDifference(XmlDifferenceNode node, XElement left, XElement right, string path, CompareOptions options)
        {
            if (left.Elements().Any() || right.Elements().Any())
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
                rightValue));
        }

        private static void AddChildDifferences(XmlDifferenceNode node, XElement left, XElement right, string path, CompareOptions options)
        {
            var leftChildren = left.Elements().ToList();
            var rightChildren = right.Elements().ToList();
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
            var value = SummarizeElement(element);
            return new XmlDifferenceNode(
                path,
                element.Name.LocalName,
                XmlDifferenceKind.Element,
                isLeftMissing ? null : value,
                isRightMissing ? null : value,
                isLeftMissing,
                isRightMissing);
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

            if (!element.Elements().Any())
            {
                builder.Append("|#=");
                builder.Append(options.NormalizeValue(NormalizeText(element.Value)));
            }

            foreach (var childSignature in element.Elements()
                         .Select(child => BuildSignature(child, options))
                         .OrderBy(signature => signature, StringComparer.Ordinal))
            {
                builder.Append("|<");
                builder.Append(childSignature);
                builder.Append('>');
            }

            return builder.ToString();
        }

        private static string SummarizeElement(XElement element)
        {
            if (!element.Elements().Any())
            {
                return NormalizeText(element.Value);
            }

            var childCount = element.Elements().Count();
            var attributeCount = element.Attributes().Count(attribute => !attribute.IsNamespaceDeclaration);
            return $"{childCount} child element(s), {attributeCount} attribute(s)";
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
    }
}
