using System;
using System.Collections.Generic;
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
        private List<ParameterDifference> _differences = new();
        private ICollectionView _differencesView;
        private readonly NotificationManager _notificationManager = new();
        private string? _leftFilePath;
        private string? _rightFilePath;
        private string _filterText = string.Empty;
        private bool _onlyShowDifferentParameters;
        private bool _ignoreCaseValues;
        private string _resultSummary = string.Empty;
        private bool _isBusy;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;

            _differencesView = CreateDifferencesView(_differences);

            BrowseLeftCommand = new RelayCommand(_ => BrowseForFile(filePath => LeftFilePath = filePath));
            BrowseRightCommand = new RelayCommand(_ => BrowseForFile(filePath => RightFilePath = filePath));
            CompareCommand = new RelayCommand(async _ => await CompareFilesAsync(), _ => CanCompareFiles());
            ExportCommand = new RelayCommand(_ => ExportResults(), _ => !IsBusy && _differencesView.Cast<ParameterDifference>().Any());

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

        public ICollectionView DifferencesView
        {
            get => _differencesView;
            private set
            {
                if (!ReferenceEquals(_differencesView, value))
                {
                    _differencesView = value;
                    OnPropertyChanged(nameof(DifferencesView));
                }
            }
        }

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
                    _differencesView.Refresh();
                    UpdateResultSummary();
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
                    _differencesView.Refresh();
                    UpdateResultSummary();
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
                    _differencesView.Refresh();
                    UpdateResultSummary();
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

                var differences = await Task.Run(() => _comparer.Compare(leftPath, rightPath).ToList());

                SetDifferences(differences);
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
            SetDifferences(Array.Empty<ParameterDifference>());
            ResultSummary = string.Empty;
        }

        private ICollectionView CreateDifferencesView(List<ParameterDifference> differences)
        {
            var view = new ListCollectionView(differences);
            view.Filter = FilterDifferences;
            view.GroupDescriptions.Add(new PropertyGroupDescription(nameof(ParameterDifference.Title)));
            return view;
        }

        private void SetDifferences(IReadOnlyCollection<ParameterDifference> differences)
        {
            _differences = differences.Count == 0
                ? new List<ParameterDifference>()
                : new List<ParameterDifference>(differences);

            DifferencesView = CreateDifferencesView(_differences);
            DifferencesView.Refresh();
            UpdateResultSummary();
            RaiseCommandStates();
        }

        private bool FilterDifferences(object obj)
        {
            if (obj is not ParameterDifference difference)
            {
                return false;
            }

            var hasDifferentValue = difference.HasDifferentValue(_ignoreCaseValues);

            if (_onlyShowDifferentParameters && !hasDifferentValue)
            {
                return false;
            }

            if (!string.IsNullOrWhiteSpace(_filterText))
            {
                return difference.Name.Contains(_filterText, StringComparison.OrdinalIgnoreCase)
                    || difference.Title.Contains(_filterText, StringComparison.OrdinalIgnoreCase)
                    || difference.DisplayName.Contains(_filterText, StringComparison.OrdinalIgnoreCase);
            }

            return true;
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
                    var differencesToExport = _differencesView.Cast<ParameterDifference>().ToList();
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
            var visibleDifferences = _differencesView.Cast<ParameterDifference>().ToList();
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
