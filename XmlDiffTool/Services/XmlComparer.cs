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
        private const string IdentityAttributeName = "Id";

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
            var node = new XmlDifferenceNode(path, leftName, XmlDifferenceKind.Element, idValue: GetIdValue(left));

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

                var leftAttribute = isLeftMissing ? null : left.Attributes().FirstOrDefault(attribute => !attribute.IsNamespaceDeclaration && options.NamesEqual(attribute.Name.LocalName, name));
                var rightAttribute = isRightMissing ? null : right.Attributes().FirstOrDefault(attribute => !attribute.IsNamespaceDeclaration && options.NamesEqual(attribute.Name.LocalName, name));
                if (!isLeftMissing
                    && !isRightMissing
                    && TryCreateDelimitedAttributeDifference(name, leftValue!, rightValue!, path, GetLineNumber(leftAttribute), GetLineNumber(rightAttribute), options, out var parameterNode))
                {
                    if (parameterNode.HasChildren)
                    {
                        node.Children.Add(parameterNode);
                    }

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
                    GetLineNumber(leftAttribute),
                    GetLineNumber(rightAttribute)));
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

            if (TryAddDelimitedParameterDifferences(node, leftValue, rightValue, path, GetLineNumber(left), GetLineNumber(right), options))
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

        private static bool TryAddDelimitedParameterDifferences(XmlDifferenceNode node, string leftValue, string rightValue, string path, int? leftLineNumber, int? rightLineNumber, CompareOptions options)
        {
            if (!TryParseDelimitedParameterTree(leftValue, options, out var leftGroup)
                || !TryParseDelimitedParameterTree(rightValue, options, out var rightGroup))
            {
                return false;
            }

            AddDelimitedGroupDifferences(node, leftGroup, rightGroup, path, leftLineNumber, rightLineNumber, options);
            return true;
        }

        private static bool TryCreateDelimitedAttributeDifference(string attributeName, string leftValue, string rightValue, string parentPath, int? leftLineNumber, int? rightLineNumber, CompareOptions options, out XmlDifferenceNode node)
        {
            var path = $"{parentPath}[@{attributeName}]";
            node = new XmlDifferenceNode(path, $"@{attributeName}", XmlDifferenceKind.Attribute, leftLineNumber: leftLineNumber, rightLineNumber: rightLineNumber);

            if (!TryParseDelimitedParameterTree(leftValue, options, out var leftGroup)
                || !TryParseDelimitedParameterTree(rightValue, options, out var rightGroup))
            {
                return false;
            }

            AddDelimitedGroupDifferences(node, leftGroup, rightGroup, path, leftLineNumber, rightLineNumber, options);
            return true;
        }

        private static void AddDelimitedGroupDifferences(XmlDifferenceNode node, DelimitedParameterGroup leftGroup, DelimitedParameterGroup rightGroup, string path, int? leftLineNumber, int? rightLineNumber, CompareOptions options)
        {
            AddDelimitedParameterRows(node, leftGroup.Parameters, rightGroup.Parameters, path, leftLineNumber, rightLineNumber, options);

            var groupNames = new HashSet<string>(leftGroup.Groups.Keys, options.NameComparer);
            groupNames.UnionWith(rightGroup.Groups.Keys);

            foreach (var groupName in groupNames.OrderBy(name => name, StringComparer.OrdinalIgnoreCase))
            {
                leftGroup.Groups.TryGetValue(groupName, out var leftChildGroup);
                rightGroup.Groups.TryGetValue(groupName, out var rightChildGroup);

                var displayName = leftChildGroup?.Name ?? rightChildGroup!.Name;
                var groupNode = new XmlDifferenceNode(
                    AppendPath(path, displayName),
                    displayName,
                    XmlDifferenceKind.Element,
                    isLeftMissing: leftChildGroup is null,
                    isRightMissing: rightChildGroup is null,
                    leftLineNumber: leftChildGroup is null ? null : leftLineNumber,
                    rightLineNumber: rightChildGroup is null ? null : rightLineNumber);

                if (leftChildGroup is null)
                {
                    AddMissingDelimitedGroupRows(groupNode, rightChildGroup!, groupNode.Path, isLeftMissing: true, lineNumber: rightLineNumber);
                }
                else if (rightChildGroup is null)
                {
                    AddMissingDelimitedGroupRows(groupNode, leftChildGroup, groupNode.Path, isRightMissing: true, lineNumber: leftLineNumber);
                }
                else
                {
                    AddDelimitedGroupDifferences(groupNode, leftChildGroup, rightChildGroup, groupNode.Path, leftLineNumber, rightLineNumber, options);
                }

                if (groupNode.HasChildren)
                {
                    node.Children.Add(groupNode);
                }
            }
        }

        private static void AddDelimitedParameterRows(XmlDifferenceNode node, Dictionary<string, DelimitedParameter> leftParameters, Dictionary<string, DelimitedParameter> rightParameters, string path, int? leftLineNumber, int? rightLineNumber, CompareOptions options)
        {
            var names = new HashSet<string>(leftParameters.Keys, options.NameComparer);
            names.UnionWith(rightParameters.Keys);

            foreach (var name in names.OrderBy(name => name, StringComparer.OrdinalIgnoreCase))
            {
                leftParameters.TryGetValue(name, out var leftParameter);
                rightParameters.TryGetValue(name, out var rightParameter);

                var isLeftMissing = leftParameter is null;
                var isRightMissing = rightParameter is null;
                if (!isLeftMissing && !isRightMissing && options.ValuesEqual(leftParameter!.Value, rightParameter!.Value))
                {
                    continue;
                }

                AddDelimitedParameterRow(node, path, leftParameter, rightParameter, isLeftMissing, isRightMissing, leftLineNumber, rightLineNumber);
            }
        }

        private static void AddDelimitedParameterRow(XmlDifferenceNode node, string path, DelimitedParameter? leftParameter, DelimitedParameter? rightParameter, bool isLeftMissing, bool isRightMissing, int? leftLineNumber, int? rightLineNumber)
        {
            var displayName = leftParameter?.Name ?? rightParameter!.Name;
            node.Children.Add(new XmlDifferenceNode(
                $"{path}[#{displayName}]",
                displayName,
                XmlDifferenceKind.Value,
                leftParameter?.Value,
                rightParameter?.Value,
                isLeftMissing,
                isRightMissing,
                isLeftMissing ? null : leftLineNumber,
                isRightMissing ? null : rightLineNumber));
        }

        private static void AddMissingDelimitedGroupRows(XmlDifferenceNode node, DelimitedParameterGroup group, string path, bool isLeftMissing = false, bool isRightMissing = false, int? lineNumber = null)
        {
            foreach (var parameter in group.Parameters.Values.OrderBy(parameter => parameter.Name, StringComparer.OrdinalIgnoreCase))
            {
                AddDelimitedParameterRow(
                    node,
                    path,
                    isRightMissing ? parameter : null,
                    isLeftMissing ? parameter : null,
                    isLeftMissing,
                    isRightMissing,
                    isRightMissing ? lineNumber : null,
                    isLeftMissing ? lineNumber : null);
            }

            foreach (var childGroup in group.Groups.Values.OrderBy(group => group.Name, StringComparer.OrdinalIgnoreCase))
            {
                var childNode = new XmlDifferenceNode(
                    AppendPath(path, childGroup.Name),
                    childGroup.Name,
                    XmlDifferenceKind.Element,
                    isLeftMissing: isLeftMissing,
                    isRightMissing: isRightMissing,
                    leftLineNumber: isRightMissing ? lineNumber : null,
                    rightLineNumber: isLeftMissing ? lineNumber : null);

                AddMissingDelimitedGroupRows(childNode, childGroup, childNode.Path, isLeftMissing, isRightMissing, lineNumber);
                node.Children.Add(childNode);
            }
        }

        private static bool TryParseDelimitedParameterTree(string value, CompareOptions options, out DelimitedParameterGroup root)
        {
            root = new DelimitedParameterGroup(string.Empty, options.NameComparer);
            if (string.IsNullOrWhiteSpace(value) || !value.Contains(';') || !value.Contains('='))
            {
                return false;
            }

            var index = 0;
            var parsedCount = 0;
            while (index < value.Length)
            {
                SkipParameterSeparators(value, ref index);

                if (index < value.Length && value[index] == '(')
                {
                    if (!TryParseDelimitedGroup(value, ref index, root, options, ref parsedCount))
                    {
                        index++;
                    }

                    continue;
                }

                if (TryParseDelimitedParameter(value, ref index, root, options))
                {
                    parsedCount++;
                    continue;
                }

                index++;
            }

            return parsedCount >= 2;
        }

        private static bool TryParseDelimitedGroup(string value, ref int index, DelimitedParameterGroup parent, CompareOptions options, ref int parsedCount)
        {
            index++;
            var nameStart = index;
            var startParsedCount = parsedCount;
            while (index < value.Length && value[index] != ':' && value[index] != ')')
            {
                index++;
            }

            if (index >= value.Length || value[index] != ':')
            {
                return false;
            }

            var groupName = value[nameStart..index].Trim();
            if (string.IsNullOrWhiteSpace(groupName))
            {
                return false;
            }

            index++;
            var group = new DelimitedParameterGroup(GetUniqueName(parent.Groups, groupName), options.NameComparer);
            parent.Groups.Add(group.Name, group);

            while (index < value.Length)
            {
                SkipParameterSeparators(value, ref index);
                if (index >= value.Length)
                {
                    break;
                }

                if (value[index] == ')')
                {
                    index++;
                    return true;
                }

                if (value[index] == '(')
                {
                    if (!TryParseDelimitedGroup(value, ref index, group, options, ref parsedCount))
                    {
                        index++;
                    }

                    continue;
                }

                if (TryParseDelimitedParameter(value, ref index, group, options))
                {
                    parsedCount++;
                    continue;
                }

                index++;
            }

            return parsedCount > startParsedCount;
        }

        private static bool TryParseDelimitedParameter(string value, ref int index, DelimitedParameterGroup group, CompareOptions options)
        {
            var nameStart = index;
            while (index < value.Length && value[index] != '=' && value[index] != ';' && value[index] != ')' && value[index] != '(')
            {
                index++;
            }

            if (index >= value.Length || value[index] != '=')
            {
                return false;
            }

            var name = value[nameStart..index].Trim();
            if (string.IsNullOrWhiteSpace(name))
            {
                return false;
            }

            index++;
            var valueStart = index;
            while (index < value.Length && value[index] != ';' && value[index] != ')' && value[index] != '(')
            {
                index++;
            }

            var parameterValue = value[valueStart..index].Trim();
            var uniqueName = GetUniqueName(group.Parameters, name);
            group.Parameters.Add(uniqueName, new DelimitedParameter(uniqueName, parameterValue));
            return true;
        }

        private static void SkipParameterSeparators(string value, ref int index)
        {
            while (index < value.Length && (char.IsWhiteSpace(value[index]) || value[index] == ';'))
            {
                index++;
            }
        }

        private static string GetUniqueName<TValue>(Dictionary<string, TValue> values, string name)
        {
            var uniqueName = name;
            var duplicateIndex = 2;
            while (values.ContainsKey(uniqueName))
            {
                uniqueName = $"{name}[{duplicateIndex}]";
                duplicateIndex++;
            }

            return uniqueName;
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
                AddIdMatchedChildDifferences(node, leftGroup, rightGroup, path, options);

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

        private static void AddIdMatchedChildDifferences(XmlDifferenceNode node, List<XElement> leftGroup, List<XElement> rightGroup, string path, CompareOptions options)
        {
            var leftById = BuildIdLookup(leftGroup, options);
            var rightById = BuildIdLookup(rightGroup, options);
            if (leftById.Count == 0 && rightById.Count == 0)
            {
                return;
            }

            var idKeys = new HashSet<string>(leftById.Keys, StringComparer.Ordinal);
            idKeys.UnionWith(rightById.Keys);

            foreach (var idKey in idKeys.OrderBy(key => key, StringComparer.Ordinal))
            {
                leftById.TryGetValue(idKey, out var leftMatches);
                rightById.TryGetValue(idKey, out var rightMatches);

                var pairedCount = Math.Min(leftMatches?.Count ?? 0, rightMatches?.Count ?? 0);
                for (var index = 0; index < pairedCount; index++)
                {
                    var left = leftMatches![index];
                    var right = rightMatches![index];
                    node.Children.AddRange(CompareElements(left, right, path, options));
                    leftGroup.Remove(left);
                    rightGroup.Remove(right);
                }

                if (leftMatches is not null)
                {
                    for (var index = pairedCount; index < leftMatches.Count; index++)
                    {
                        var left = leftMatches[index];
                        node.Children.Add(CreateMissingElementNode(left, path, isRightMissing: true));
                        leftGroup.Remove(left);
                    }
                }

                if (rightMatches is not null)
                {
                    for (var index = pairedCount; index < rightMatches.Count; index++)
                    {
                        var right = rightMatches[index];
                        node.Children.Add(CreateMissingElementNode(right, path, isLeftMissing: true));
                        rightGroup.Remove(right);
                    }
                }
            }
        }

        private static Dictionary<string, List<XElement>> BuildIdLookup(IEnumerable<XElement> elements, CompareOptions options)
        {
            var lookup = new Dictionary<string, List<XElement>>(StringComparer.Ordinal);
            foreach (var element in elements)
            {
                var idKey = GetIdKey(element, options);
                if (idKey is null)
                {
                    continue;
                }

                if (!lookup.TryGetValue(idKey, out var matches))
                {
                    matches = new List<XElement>();
                    lookup.Add(idKey, matches);
                }

                matches.Add(element);
            }

            return lookup;
        }

        private static string? GetIdKey(XElement element, CompareOptions options)
        {
            var idAttribute = GetIdAttribute(element);

            return idAttribute is null ? null : options.NormalizeValue(idAttribute.Value);
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
                isRightMissing ? null : GetLineNumber(element),
                GetIdValue(element));

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
                rightLineNumber: isRightMissing ? null : GetLineNumber(element),
                idValue: GetIdValue(element));

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

        private static XAttribute? GetIdAttribute(XElement element)
        {
            return element.Attributes()
                .FirstOrDefault(attribute => !attribute.IsNamespaceDeclaration
                                             && string.Equals(attribute.Name.LocalName, IdentityAttributeName, StringComparison.OrdinalIgnoreCase));
        }

        private static string? GetIdValue(XElement element)
        {
            return GetIdAttribute(element)?.Value;
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

        private sealed class DelimitedParameter
        {
            public DelimitedParameter(string name, string value)
            {
                Name = name;
                Value = value;
            }

            public string Name { get; }

            public string Value { get; }
        }

        private sealed class DelimitedParameterGroup
        {
            public DelimitedParameterGroup(string name, StringComparer comparer)
            {
                Name = name;
                Parameters = new Dictionary<string, DelimitedParameter>(comparer);
                Groups = new Dictionary<string, DelimitedParameterGroup>(comparer);
            }

            public string Name { get; }

            public Dictionary<string, DelimitedParameter> Parameters { get; }

            public Dictionary<string, DelimitedParameterGroup> Groups { get; }
        }
    }
}
