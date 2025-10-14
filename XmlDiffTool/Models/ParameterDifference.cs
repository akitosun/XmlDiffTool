using System;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;

namespace XmlDiffTool.Models
{
    public class ParameterDifference : INotifyPropertyChanged
    {
        private readonly ObservableCollection<ParameterDifference> _children = new();
        private readonly ReadOnlyObservableCollection<ParameterDifference> _readOnlyChildren;
        private string? _leftValue;
        private string? _rightValue;
        private bool _isExpanded = true;
        private bool _isVisible = true;

        public ParameterDifference(string name, string displayName, ParameterDifference? parent)
        {
            Name = name;
            DisplayName = displayName;
            Parent = parent;

            _readOnlyChildren = new ReadOnlyObservableCollection<ParameterDifference>(_children);
            _children.CollectionChanged += (_, _) => OnPropertyChanged(nameof(HasChildren));
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        public string Name { get; }

        public string DisplayName { get; }

        public ParameterDifference? Parent { get; }

        public ReadOnlyObservableCollection<ParameterDifference> Children => _readOnlyChildren;

        public bool HasChildren => _children.Any();

        public int Level => Parent is null ? 0 : Parent.Level + 1;

        public string? LeftValue
        {
            get => _leftValue;
            private set
            {
                if (_leftValue != value)
                {
                    _leftValue = value;
                    OnPropertyChanged(nameof(LeftValue));
                    OnPropertyChanged(nameof(IsLeftMissing));
                    OnPropertyChanged(nameof(HasMissingValue));
                }
            }
        }

        public string? RightValue
        {
            get => _rightValue;
            private set
            {
                if (_rightValue != value)
                {
                    _rightValue = value;
                    OnPropertyChanged(nameof(RightValue));
                    OnPropertyChanged(nameof(IsRightMissing));
                    OnPropertyChanged(nameof(HasMissingValue));
                }
            }
        }

        public bool IsExpanded
        {
            get => _isExpanded;
            set
            {
                if (_isExpanded != value)
                {
                    _isExpanded = value;
                    OnPropertyChanged(nameof(IsExpanded));
                }
            }
        }

        public bool IsVisible
        {
            get => _isVisible;
            set
            {
                if (_isVisible != value)
                {
                    _isVisible = value;
                    OnPropertyChanged(nameof(IsVisible));
                }
            }
        }

        public bool IsLeftMissing => !HasChildren && string.IsNullOrWhiteSpace(LeftValue);

        public bool IsRightMissing => !HasChildren && string.IsNullOrWhiteSpace(RightValue);

        public bool HasMissingValue => IsLeftMissing || IsRightMissing;

        public bool HasDifferentValue(bool ignoreCase)
        {
            if (HasChildren)
            {
                return false;
            }

            var comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            return !string.Equals(LeftValue, RightValue, comparison);
        }

        internal void SetValues(string? leftValue, string? rightValue)
        {
            LeftValue = leftValue;
            RightValue = rightValue;
        }

        internal ParameterDifference GetOrCreateChild(string fullPath, string displayName)
        {
            var existing = _children.FirstOrDefault(child => child.DisplayName == displayName);
            if (existing is not null)
            {
                return existing;
            }

            var child = new ParameterDifference(fullPath, displayName, this);
            _children.Add(child);
            return child;
        }

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
