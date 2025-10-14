using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Win32;
using Notifications.Wpf;
using XmlDiffTool.Infrastructure;
using XmlDiffTool.Models;
using XmlDiffTool.Services;

namespace XmlDiffTool
{
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private readonly XmlComparer _comparer = new();
        private readonly ObservableCollection<ParameterDifference> _differences = new();
        private readonly HashSet<string> _userCollapsedNodes = new();
        private readonly List<ParameterDifference> _differenceTree = new();
        private readonly ICollectionView _differencesView;
        private readonly NotificationManager _notificationManager = new();
        private string? _leftFilePath;
        private string? _rightFilePath;
        private string _filterText = string.Empty;
        private bool _onlyShowDifferentParameters;
        private bool _ignoreCaseValues;
        private string _resultSummary = string.Empty;
        private bool _isBusy;
        private bool _isApplyingFilters;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;

            _differencesView = CollectionViewSource.GetDefaultView(_differences);

            BrowseLeftCommand = new RelayCommand(_ => BrowseForFile(filePath => LeftFilePath = filePath));
            BrowseRightCommand = new RelayCommand(_ => BrowseForFile(filePath => RightFilePath = filePath));
            CompareCommand = new RelayCommand(async _ => await CompareFilesAsync(), _ => CanCompareFiles());
            ExportCommand = new RelayCommand(_ => ExportResults(), _ => !IsBusy && GetVisibleLeaves().Any());

            UpdateResultSummary();
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        public ICommand BrowseLeftCommand { get; }

        public ICommand BrowseRightCommand { get; }

        public ICommand CompareCommand { get; }

        public ICommand ExportCommand { get; }

        public bool IsBusy
        {
            get => _isBusy;
            private set
            {
                if (_isBusy != value)
                {
                    _isBusy = value;
                    OnPropertyChanged(nameof(IsBusy));
                    RaiseCommandStates();
                }
            }
        }

        public ICollectionView DifferencesView => _differencesView;

        public string? LeftFilePath
        {
            get => _leftFilePath;
            set
            {
                if (_leftFilePath != value)
                {
                    _leftFilePath = value;
                    OnPropertyChanged(nameof(LeftFilePath));
                    RaiseCommandStates();
                }
            }
        }

        public string? RightFilePath
        {
            get => _rightFilePath;
            set
            {
                if (_rightFilePath != value)
                {
                    _rightFilePath = value;
                    OnPropertyChanged(nameof(RightFilePath));
                    RaiseCommandStates();
                }
            }
        }

        public string FilterText
        {
            get => _filterText;
            set
            {
                if (_filterText != value)
                {
                    _filterText = value;
                    OnPropertyChanged(nameof(FilterText));
                    ApplyFilters(true);
                }
            }
        }

        public bool OnlyShowDifferentParameters
        {
            get => _onlyShowDifferentParameters;
            set
            {
                if (_onlyShowDifferentParameters != value)
                {
                    _onlyShowDifferentParameters = value;
                    OnPropertyChanged(nameof(OnlyShowDifferentParameters));
                    ApplyFilters(true);
                }
            }
        }

        public bool IgnoreCaseValues
        {
            get => _ignoreCaseValues;
            set
            {
                if (_ignoreCaseValues != value)
                {
                    _ignoreCaseValues = value;
                    OnPropertyChanged(nameof(IgnoreCaseValues));
                    ApplyFilters(true);
                }
            }
        }

        public string ResultSummary
        {
            get => _resultSummary;
            private set
            {
                if (_resultSummary != value)
                {
                    _resultSummary = value;
                    OnPropertyChanged(nameof(ResultSummary));
                }
            }
        }

        private void BrowseForFile(Action<string> onFileSelected)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*"
            };

            if (dialog.ShowDialog() == true)
            {
                onFileSelected(dialog.FileName);
            }
        }

        private bool CanCompareFiles()
        {
            return !IsBusy && File.Exists(_leftFilePath) && File.Exists(_rightFilePath);
        }

        private async Task CompareFilesAsync()
        {
            try
            {
                IsBusy = true;
                ClearResults();

                var leftPath = _leftFilePath!;
                var rightPath = _rightFilePath!;

                var differences = await Task.Run(() => _comparer.Compare(leftPath, rightPath));

                foreach (var difference in differences)
                {
                    AttachDifferenceHandlers(difference);
                    _differenceTree.Add(difference);
                }

                ApplyFilters(true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Failed to compare XML files.\n{ex.Message}", "Comparison Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
            }
        }

        private void ClearResults()
        {
            _differences.Clear();
            _differenceTree.Clear();
            _userCollapsedNodes.Clear();
            _differencesView.Refresh();
            ResultSummary = string.Empty;
        }

        private void AttachDifferenceHandlers(ParameterDifference difference)
        {
            difference.PropertyChanged += ParameterDifferenceOnPropertyChanged;
            foreach (var child in difference.Children)
            {
                AttachDifferenceHandlers(child);
            }
        }

        private void ParameterDifferenceOnPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (_isApplyingFilters)
            {
                return;
            }

            if (sender is not ParameterDifference difference)
            {
                return;
            }

            if (e.PropertyName == nameof(ParameterDifference.IsExpanded))
            {
                if (difference.IsExpanded)
                {
                    _userCollapsedNodes.Remove(difference.Name);
                }
                else
                {
                    _userCollapsedNodes.Add(difference.Name);
                }

                RefreshVisibleDifferences();
            }
        }

        private void ApplyFilters(bool forceExpand)
        {
            _isApplyingFilters = true;
            try
            {
                foreach (var root in _differenceTree)
                {
                    UpdateVisibility(root, forceExpand);
                }
            }
            finally
            {
                _isApplyingFilters = false;
            }

            RefreshVisibleDifferences();
        }

        private bool UpdateVisibility(ParameterDifference node, bool forceExpand)
        {
            var childVisible = false;

            foreach (var child in node.Children)
            {
                if (UpdateVisibility(child, forceExpand))
                {
                    childVisible = true;
                }
            }

            var matchesSelf = MatchesNode(node);
            var isVisible = matchesSelf || childVisible;

            node.IsVisible = isVisible;

            if (isVisible && node.HasChildren && forceExpand && !_userCollapsedNodes.Contains(node.Name))
            {
                node.IsExpanded = true;
            }

            return isVisible;
        }

        private bool MatchesNode(ParameterDifference node)
        {
            if (!string.IsNullOrWhiteSpace(_filterText) &&
                !node.Name.Contains(_filterText, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (!node.HasChildren)
            {
                if (_onlyShowDifferentParameters && !node.HasDifferentValue(_ignoreCaseValues))
                {
                    return false;
                }

                return true;
            }

            return !string.IsNullOrWhiteSpace(_filterText);
        }

        private void RefreshVisibleDifferences()
        {
            _differences.Clear();

            foreach (var root in _differenceTree)
            {
                foreach (var difference in FlattenVisible(root))
                {
                    _differences.Add(difference);
                }
            }

            _differencesView.Refresh();
            UpdateResultSummary();
        }

        private IEnumerable<ParameterDifference> FlattenVisible(ParameterDifference node)
        {
            if (!node.IsVisible)
            {
                yield break;
            }

            yield return node;

            if (!node.IsExpanded)
            {
                yield break;
            }

            foreach (var child in node.Children)
            {
                foreach (var descendant in FlattenVisible(child))
                {
                    yield return descendant;
                }
            }
        }

        private IEnumerable<ParameterDifference> GetVisibleLeaves()
        {
            foreach (var root in _differenceTree)
            {
                foreach (var difference in GetVisibleLeaves(root))
                {
                    yield return difference;
                }
            }
        }

        private IEnumerable<ParameterDifference> GetVisibleLeaves(ParameterDifference node)
        {
            if (!node.IsVisible)
            {
                yield break;
            }

            if (!node.HasChildren)
            {
                yield return node;
                yield break;
            }

            foreach (var child in node.Children)
            {
                foreach (var difference in GetVisibleLeaves(child))
                {
                    yield return difference;
                }
            }
        }

        private void ExportResults()
        {
            var dialog = new SaveFileDialog
            {
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                FileName = "XmlDiffResults.xlsx"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    var differencesToExport = GetVisibleLeaves().ToList();
                    if (!differencesToExport.Any())
                    {
                        MessageBox.Show(this, "There are no results to export.", "Export", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }

                    using var document = SpreadsheetDocument.Create(dialog.FileName, SpreadsheetDocumentType.Workbook);
                    var workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();

                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());

                    var sheets = document.WorkbookPart!.Workbook.AppendChild(new Sheets());
                    var sheet = new Sheet
                    {
                        Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "Differences"
                    };
                    sheets.Append(sheet);

                    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                    AppendRow(sheetData, "Name", "Left XML Value", "Right XML Value");

                    foreach (var difference in differencesToExport)
                    {
                        AppendRow(sheetData, difference.Name, difference.LeftValue ?? string.Empty, difference.RightValue ?? string.Empty);
                    }

                    worksheetPart.Worksheet.Save();
                    workbookPart.Workbook.Save();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, $"Failed to export results.\n{ex.Message}", "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private static void AppendRow(SheetData sheetData, params string[] values)
        {
            var row = new Row();
            foreach (var value in values)
            {
                var cell = new Cell
                {
                    DataType = CellValues.String,
                    CellValue = new CellValue(value)
                };
                row.Append(cell);
            }

            sheetData.Append(row);
        }

        private void UpdateResultSummary()
        {
            var visibleDifferences = GetVisibleLeaves().ToList();
            var totalCount = visibleDifferences.Count;
            var leftMissingCount = visibleDifferences.Count(d => d.IsLeftMissing);
            var rightMissingCount = visibleDifferences.Count(d => d.IsRightMissing);

            ResultSummary = $"Result: All {totalCount} parameters difference, and Left miss {leftMissingCount} parameters,Right miss {rightMissingCount} parameters.";

            (ExportCommand as RelayCommand)?.RaiseCanExecuteChanged();
        }

        private void RaiseCommandStates()
        {
            (CompareCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (ExportCommand as RelayCommand)?.RaiseCanExecuteChanged();
        }

        private void DifferencesDataGrid_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (sender is not DataGrid dataGrid)
            {
                return;
            }

            if (dataGrid.SelectedItem is not ParameterDifference difference || string.IsNullOrWhiteSpace(difference.Name))
            {
                return;
            }

            try
            {
                Clipboard.SetText(difference.Name);
            }
            catch
            {
                return;
            }

            _notificationManager.Show(new NotificationContent
            {
                Title = "Copied",
                Message = $"{difference.Name} is copied.",
                Type = NotificationType.Information
            });
        }

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
