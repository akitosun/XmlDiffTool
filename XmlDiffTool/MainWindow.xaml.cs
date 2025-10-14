using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Win32;
using XmlDiffTool.Infrastructure;
using XmlDiffTool.Models;
using XmlDiffTool.Services;

namespace XmlDiffTool
{
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private readonly XmlComparer _comparer = new();
        private readonly ObservableCollection<ParameterDifference> _differences = new();
        private readonly ICollectionView _differencesView;
        private string? _leftFilePath;
        private string? _rightFilePath;
        private string _filterText = string.Empty;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;

            _differencesView = CollectionViewSource.GetDefaultView(_differences);
            _differencesView.Filter = FilterDifferences;

            BrowseLeftCommand = new RelayCommand(_ => BrowseForFile(filePath => LeftFilePath = filePath));
            BrowseRightCommand = new RelayCommand(_ => BrowseForFile(filePath => RightFilePath = filePath));
            CompareCommand = new RelayCommand(_ => CompareFiles(), _ => CanCompareFiles());
            ExportCommand = new RelayCommand(_ => ExportResults(), _ => _differences.Any());
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        public ICommand BrowseLeftCommand { get; }

        public ICommand BrowseRightCommand { get; }

        public ICommand CompareCommand { get; }

        public ICommand ExportCommand { get; }

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
                    _differencesView.Refresh();
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
            return File.Exists(_leftFilePath) && File.Exists(_rightFilePath);
        }

        private void CompareFiles()
        {
            try
            {
                _differences.Clear();
                foreach (var difference in _comparer.Compare(_leftFilePath!, _rightFilePath!))
                {
                    _differences.Add(difference);
                }

                _differencesView.Refresh();
                RaiseCommandStates();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Failed to compare XML files.\n{ex.Message}", "Comparison Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool FilterDifferences(object obj)
        {
            if (obj is not ParameterDifference difference)
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(_filterText))
            {
                return true;
            }

            return difference.Name.Contains(_filterText, StringComparison.OrdinalIgnoreCase);
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

                    foreach (var difference in _differences)
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

        private void RaiseCommandStates()
        {
            (CompareCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (ExportCommand as RelayCommand)?.RaiseCanExecuteChanged();
        }

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
