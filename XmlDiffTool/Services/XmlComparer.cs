using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using XmlDiffTool.Models;

namespace XmlDiffTool.Services
{
    public class XmlComparer
    {
        public IReadOnlyCollection<ParameterDifference> Compare(string leftPath, string rightPath)
        {
            var leftParameters = LoadParameters(leftPath);
            var rightParameters = LoadParameters(rightPath);

            var keys = new HashSet<string>(leftParameters.Keys);
            keys.UnionWith(rightParameters.Keys);

            return keys
                .OrderBy(k => k)
                .Select(key =>
                {
                    var (title, displayName) = ParseKey(key);

                    return new ParameterDifference(
                        key,
                        title,
                        displayName,
                        leftParameters.TryGetValue(key, out var leftValue) ? leftValue : null,
                        rightParameters.TryGetValue(key, out var rightValue) ? rightValue : null);
                })
                .ToList();
        }

        private static (string Title, string DisplayName) ParseKey(string key)
        {
            var attributeMarkerIndex = key.LastIndexOf("[@", System.StringComparison.Ordinal);
            if (attributeMarkerIndex >= 0)
            {
                var tagPath = key[..attributeMarkerIndex];
                var lastSlashIndex = tagPath.LastIndexOf('/');
                var tagSegment = lastSlashIndex >= 0 ? tagPath[(lastSlashIndex + 1)..] : tagPath;
                var tagNameEnd = tagSegment.IndexOf('[');
                var tagName = tagNameEnd >= 0 ? tagSegment[..tagNameEnd] : tagSegment;

                var attributeStart = attributeMarkerIndex + 2;
                var attributeEnd = key.IndexOf(']', attributeStart);
                var attributeName = attributeEnd > attributeStart
                    ? key[attributeStart..attributeEnd]
                    : key[attributeStart..];

                return (string.IsNullOrWhiteSpace(tagName) ? "Unknown" : tagName, attributeName);
            }

            var lastPathSeparatorIndex = key.LastIndexOf('/');
            var rawSegment = lastPathSeparatorIndex >= 0 ? key[(lastPathSeparatorIndex + 1)..] : key;
            var valueMarkerIndex = rawSegment.IndexOf("[#", System.StringComparison.Ordinal);
            var tagSegment = valueMarkerIndex >= 0 ? rawSegment[..valueMarkerIndex] : rawSegment;
            var tagNameEnd = tagSegment.IndexOf('[');
            var tagName = tagNameEnd >= 0 ? tagSegment[..tagNameEnd] : tagSegment;

            return (string.IsNullOrWhiteSpace(tagName) ? "Unknown" : tagName, "Value");
        }

        private static Dictionary<string, string> LoadParameters(string path)
        {
            var document = XDocument.Load(path);
            var result = new Dictionary<string, string>();
            if (document.Root is not null)
            {
                TraverseElement(document.Root, document.Root.Name.LocalName, result);
            }

            return result;
        }

        private static void TraverseElement(XElement element, string currentPath, IDictionary<string, string> values)
        {
            foreach (var attribute in element.Attributes())
            {
                var attributePath = $"{currentPath}[@{attribute.Name.LocalName}]";
                values[attributePath] = attribute.Value;
            }

            if (!element.Elements().Any())
            {
                var valuePath = $"{currentPath}[#text]";
                values[valuePath] = element.Value;
            }

            var childGroups = element.Elements()
                .GroupBy(e => e.Name.LocalName)
                .ToDictionary(g => g.Key, g => g.ToList());

            foreach (var group in childGroups)
            {
                for (var index = 0; index < group.Value.Count; index++)
                {
                    var child = group.Value[index];
                    var childPath = $"{currentPath}/{child.Name.LocalName}[{index}]";
                    TraverseElement(child, childPath, values);
                }
            }
        }
    }
}
