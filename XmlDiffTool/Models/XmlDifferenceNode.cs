using System.Collections.Generic;

namespace XmlDiffTool.Models
{
    public enum XmlDifferenceKind
    {
        Element,
        Attribute,
        Value
    }

    public class XmlDifferenceNode
    {
        public XmlDifferenceNode(string path, string name, XmlDifferenceKind kind, string? leftValue = null, string? rightValue = null, bool isLeftMissing = false, bool isRightMissing = false)
        {
            Path = path;
            Name = name;
            Kind = kind;
            LeftValue = leftValue;
            RightValue = rightValue;
            IsLeftMissing = isLeftMissing;
            IsRightMissing = isRightMissing;
        }

        public string Path { get; }

        public string Name { get; }

        public XmlDifferenceKind Kind { get; }

        public string? LeftValue { get; }

        public string? RightValue { get; }

        public bool IsLeftMissing { get; }

        public bool IsRightMissing { get; }

        public List<XmlDifferenceNode> Children { get; } = new();

        public bool HasChildren => Children.Count > 0;
    }
}
