using System;

namespace XmlDiffTool.Models
{
    public class ParameterDifference
    {
        public ParameterDifference(string name, string title, string displayName, string? leftValue, string? rightValue, bool isLeftMissing = false, bool isRightMissing = false)
        {
            Name = name;
            Title = title;
            DisplayName = displayName;
            LeftValue = leftValue;
            RightValue = rightValue;
            IsLeftMissing = isLeftMissing;
            IsRightMissing = isRightMissing;
        }

        public string Name { get; }

        public string Title { get; }

        public string DisplayName { get; }

        public string? LeftValue { get; }

        public string? RightValue { get; }

        public bool IsLeftMissing { get; }

        public bool IsRightMissing { get; }

        public bool HasMissingValue => IsLeftMissing || IsRightMissing;

        public bool HasDifferentValue(bool ignoreCase)
        {
            var comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            return !string.Equals(LeftValue, RightValue, comparison);
        }
    }
}
