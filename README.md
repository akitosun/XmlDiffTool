# XML Diff Tool

XML Diff Tool is a WPF desktop application for Windows that compares two XML files and exports the differences as a reusable HTML report. It compares XML as a tree, can ignore case, ignores list ordering for matching child elements, and only reports nodes whose attributes, values, or presence differ.

## Features

- **Side-by-side XML report** - Load a left and right XML file and generate an asynchronous HTML report for element, attribute, and value differences.
- **Order-insensitive matching** - Repeated child elements are matched by normalized content, so list ordering does not create false differences.
- **Report filtering** - The generated HTML report can hide left-only or right-only elements from the report toolbar.
- **HTML report export** - Save a Bootstrap-styled side-by-side HTML report. Shared CSS and JS assets are written under `%AppData%\XmlDiffTool\ReportAssets` and reused by later reports.

## Prerequisites

- Windows 7 or later (required for WPF)
- [.NET 6 SDK](https://dotnet.microsoft.com/en-us/download/dotnet/6.0) with Windows desktop support
- (Optional) [Visual Studio 2022](https://visualstudio.microsoft.com/) with the ".NET desktop development" workload for an IDE experience

The project uses WPF and the .NET base class libraries; no additional NuGet packages are required.

## Getting Started

Clone the repository and restore dependencies:

```bash
git clone <repository-url>
cd XmlDiffTool/XmlDiffTool
dotnet restore
```

### Running from the command line

```bash
dotnet run
```

The app launches a WPF window titled **XML Diff Tool**.

### Running from Visual Studio

1. Open `XmlDiffTool.sln` in Visual Studio 2022 or newer.
2. Set `XmlDiffTool` as the startup project.
3. Press <kbd>F5</kbd> to build and run.

## Using the Application

1. Click **Browse...** on the left and right sides to choose two XML files.
2. Enable **Ignore case for tags, attributes, and values** if matching should be case-insensitive.
3. Press **Generate HTML Report**, choose a save location, and open the generated report when prompted.
4. Use the report toolbar to show or hide left-only and right-only elements.

The WPF window only handles file selection, compare options, and report generation. Difference review happens in the generated HTML report.

## How It Works

The comparer loads each XML document as a tree. For repeated child elements with the same name, exact normalized matches are removed first so list ordering is ignored. Remaining elements are compared by attributes, leaf values, and child differences. Nodes whose attributes and values are all equal are not included in the report.

## License

This project is licensed under the [MIT License](LICENSE).
