using System.Collections.Generic;
using System.Linq;

namespace XmlDiffTool.Models
{
    public class XmlComparisonResult
    {
        public XmlComparisonResult(IReadOnlyCollection<XmlDifferenceNode> roots)
        {
            Roots = roots;
            Differences = Flatten(roots).ToList();
        }

        public IReadOnlyCollection<XmlDifferenceNode> Roots { get; }

        public IReadOnlyCollection<ParameterDifference> Differences { get; }

        private static IEnumerable<ParameterDifference> Flatten(IEnumerable<XmlDifferenceNode> nodes)
        {
            foreach (var node in nodes)
            {
                if (node.Kind != XmlDifferenceKind.Element || node.IsLeftMissing || node.IsRightMissing || !node.HasChildren)
                {
                    yield return new ParameterDifference(
                        node.Path,
                        GetTitle(node.Path),
                        node.Name,
                        node.LeftValue,
                        node.RightValue,
                        node.IsLeftMissing,
                        node.IsRightMissing,
                        node.LeftLineNumber,
                        node.RightLineNumber);
                }

                foreach (var child in Flatten(node.Children))
                {
                    yield return child;
                }
            }
        }

        private static string GetTitle(string path)
        {
            var attributeMarkerIndex = path.LastIndexOf("[@", System.StringComparison.Ordinal);
            if (attributeMarkerIndex >= 0)
            {
                return GetTagName(path[..attributeMarkerIndex]);
            }

            var valueMarkerIndex = path.LastIndexOf("[#value]", System.StringComparison.Ordinal);
            if (valueMarkerIndex >= 0)
            {
                return GetTagName(path[..valueMarkerIndex]);
            }

            return GetTagName(path);
        }

        private static string GetTagName(string path)
        {
            var lastSlashIndex = path.LastIndexOf('/');
            var tagSegment = lastSlashIndex >= 0 ? path[(lastSlashIndex + 1)..] : path;
            var tagNameEnd = tagSegment.IndexOf('[');
            var tagName = tagNameEnd >= 0 ? tagSegment[..tagNameEnd] : tagSegment;

            return string.IsNullOrWhiteSpace(tagName) ? "Unknown" : tagName;
        }
    }
}
