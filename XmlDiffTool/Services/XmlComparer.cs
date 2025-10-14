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

            var roots = new List<ParameterDifference>();

            foreach (var key in keys.OrderBy(k => k))
            {
                AddDifferenceNode(
                    roots,
                    key,
                    leftParameters.TryGetValue(key, out var leftValue) ? leftValue : null,
                    rightParameters.TryGetValue(key, out var rightValue) ? rightValue : null);
            }

            return roots;
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
                var attributePath = $"{currentPath}/@{attribute.Name.LocalName}";
                values[attributePath] = attribute.Value;
            }

            var childGroups = element.Elements()
                .GroupBy(e => e.Name.LocalName)
                .ToDictionary(g => g.Key, g => g.ToList());

            if (!element.HasElements && !string.IsNullOrWhiteSpace(element.Value))
            {
                values[currentPath] = element.Value.Trim();
            }

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

        private static void AddDifferenceNode(ICollection<ParameterDifference> roots, string path, string? leftValue, string? rightValue)
        {
            var segments = path.Split('/');
            ParameterDifference? current = null;

            foreach (var segment in segments)
            {
                var fullPath = current is null ? segment : $"{current.Name}/{segment}";
                if (current is null)
                {
                    current = roots.FirstOrDefault(d => d.DisplayName == segment);
                    if (current is null)
                    {
                        current = new ParameterDifference(fullPath, segment, null);
                        roots.Add(current);
                    }
                }
                else
                {
                    current = current.GetOrCreateChild(fullPath, segment);
                }
            }

            current?.SetValues(leftValue, rightValue);
        }
    }
}
