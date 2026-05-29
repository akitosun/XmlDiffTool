# XML Diff Tool

XML Diff Tool is a WPF desktop application for Windows that compares two XML files and exports the differences as a reusable HTML report. It compares XML as structured data, aligns repeated items intelligently, expands embedded parameter strings, and reports only the meaningful differences.

## Features

- **Side-by-side XML reports** - Compare a left and right XML file and generate an HTML report for element, attribute, value, and presence differences.
- **Case-insensitive comparison option** - Match tag names, attribute names, and values without case sensitivity when the option is enabled.
- **Order-insensitive repeated elements** - Match repeated child elements by normalized content so reordered lists stay readable.
- **Id-based item matching** - When repeated items have an `Id` attribute, regardless of attribute-name casing, items with the same Id are compared with each other.
- **Id badges in reports** - Show each compared item's Id as a blue badge beside the node heading, making list item differences easy to locate.
- **Difference count badges** - Show each report section's difference count as a clear `diff: n` badge.
- **Attribute comparison** - Report changed, missing, left-only, and right-only attributes.
- **Leaf value comparison** - Report changed text content with source line numbers when available.
- **Embedded XML comparison** - Parse XML stored inside text values and compare it as child XML instead of opaque text.
- **Delimited parameter string comparison** - Expand strings such as `(MAIN@Default:key1=0;key2=1;(CH1@CH:key1=0;))` into group nodes and key/value rows.
- **Nested parameter groups** - Display delimited-string groups as report nodes and show keys as properties inside each group.
- **Missing item reporting** - Clearly show elements, attributes, values, or parameter keys that exist only on one side.
- **Report filtering** - Hide or show left-only and right-only differences from the generated report toolbar.
- **Line number toggle** - Show or hide XML source line numbers in the report.
- **HTML report export** - Save a Bootstrap-styled report. Shared CSS and JS assets are written under `%AppData%\XmlDiffTool\ReportAssets` and reused by later reports.

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

### Building an executable with Build.bat

Run the included batch file from the repository root:

```bat
Build.bat
```

The script publishes the WPF app in Release mode and writes the executable to:

```text
bin\XmlDiffTool.exe
```

You can run that executable directly without opening Visual Studio.

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

The comparer loads each XML document as a tree. For repeated child elements with the same name, exact normalized matches are removed first so list ordering is ignored. Remaining items with an `Id` attribute are matched by Id, and the rest are matched by normalized structure. Attributes, leaf values, embedded XML, child nodes, and delimited parameter strings are then compared. Nodes whose attributes, values, parameters, and children are all equal are not included in the report.

## License

This project is licensed under the [MIT License](LICENSE).
