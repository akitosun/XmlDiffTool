using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;
using XmlDiffTool.Infrastructure;
using XmlDiffTool.Services;

namespace XmlDiffTool
{
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private readonly XmlComparer _comparer = new();
        private readonly HtmlReportService _htmlReportService = new();
        private string? _leftFilePath;
        private string? _rightFilePath;
        private bool _ignoreCaseValues = true;
        private string _resultSummary = "Select two XML files, then generate an HTML report.";
        private bool _isBusy;
        private int _progressPercentage;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;

            BrowseLeftCommand = new RelayCommand(_ => BrowseForFile(filePath => LeftFilePath = filePath));
            BrowseRightCommand = new RelayCommand(_ => BrowseForFile(filePath => RightFilePath = filePath));
            CompareCommand = new RelayCommand(async _ => await GenerateReportAsync(), _ => CanGenerateReport());
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        public ICommand BrowseLeftCommand { get; }

        public ICommand BrowseRightCommand { get; }

        public ICommand CompareCommand { get; }

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

        public bool IgnoreCaseValues
        {
            get => _ignoreCaseValues;
            set
            {
                if (_ignoreCaseValues != value)
                {
                    _ignoreCaseValues = value;
                    OnPropertyChanged(nameof(IgnoreCaseValues));
                }
            }
        }

        public int ProgressPercentage
        {
            get => _progressPercentage;
            private set
            {
                if (_progressPercentage != value)
                {
                    _progressPercentage = value;
                    OnPropertyChanged(nameof(ProgressPercentage));
                    OnPropertyChanged(nameof(ProgressText));
                }
            }
        }

        public string ProgressText => $"Generating report... {ProgressPercentage}%";

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

        private bool CanGenerateReport()
        {
            return !IsBusy && File.Exists(_leftFilePath) && File.Exists(_rightFilePath);
        }

        private async Task GenerateReportAsync()
        {
            var dialog = new SaveFileDialog
            {
                Filter = "HTML Report (*.html)|*.html",
                FileName = BuildReportFileName(_leftFilePath!, _rightFilePath!)
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            try
            {
                IsBusy = true;
                ProgressPercentage = 0;
                ResultSummary = "Comparing XML files...";

                var leftPath = _leftFilePath!;
                var rightPath = _rightFilePath!;
                var progress = new Progress<int>(value => ProgressPercentage = value);

                var comparisonResult = await Task.Run(() => _comparer.Compare(leftPath, rightPath, IgnoreCaseValues, progress));
                _htmlReportService.SaveReport(dialog.FileName, comparisonResult.Roots, LeftFilePath, RightFilePath, IgnoreCaseValues);

                var count = CountNodes(comparisonResult.Roots);
                ResultSummary = $"Report generated. Differences: {count}.";

                if (MessageBox.Show(this, "HTML report has been created. Open it now?", "Report Complete", MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
                {
                    Process.Start(new ProcessStartInfo(dialog.FileName)
                    {
                        UseShellExecute = true
                    });
                }
            }
            catch (Exception ex)
            {
                ResultSummary = "Report generation failed.";
                MessageBox.Show(this, $"Failed to generate HTML report.\n{ex.Message}", "Report Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
                ProgressPercentage = 0;
            }
        }

        private static int CountNodes(System.Collections.Generic.IEnumerable<Models.XmlDifferenceNode> nodes)
        {
            return nodes.Sum(node => 1 + CountNodes(node.Children));
        }

        private static string BuildReportFileName(string leftFilePath, string rightFilePath)
        {
            var leftName = SanitizeFileName(Path.GetFileNameWithoutExtension(leftFilePath));
            var rightName = SanitizeFileName(Path.GetFileNameWithoutExtension(rightFilePath));
            return $"{leftName}_{rightName}_diff.html";
        }

        private static string SanitizeFileName(string value)
        {
            foreach (var invalidChar in Path.GetInvalidFileNameChars())
            {
                value = value.Replace(invalidChar, '_');
            }

            return string.IsNullOrWhiteSpace(value) ? "xml" : value;
        }

        private void RaiseCommandStates()
        {
            (CompareCommand as RelayCommand)?.RaiseCanExecuteChanged();
        }

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
