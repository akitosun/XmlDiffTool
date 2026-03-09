using System;

namespace XmlDiffTool.Models
{
    public class ParameterDifference
    {
        public ParameterDifference(string name, string title, string displayName, string? leftValue, string? rightValue)
        {
            Name = name;
            Title = title;
            DisplayName = displayName;
            LeftValue = leftValue;
            RightValue = rightValue;
        }

        public string Name { get; }

        public string Title { get; }

        public string DisplayName { get; }

        public string? LeftValue { get; }

        public string? RightValue { get; }

        public bool IsLeftMissing => string.IsNullOrWhiteSpace(LeftValue);

        public bool IsRightMissing => string.IsNullOrWhiteSpace(RightValue);

        public bool HasMissingValue => IsLeftMissing || IsRightMissing;

        public bool HasDifferentValue(bool ignoreCase)
        {
            var comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            return !string.Equals(LeftValue, RightValue, comparison);
        }
    }
}
