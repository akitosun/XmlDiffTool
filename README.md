# XML Diff Tool

XML Diff Tool is a WPF desktop application for Windows that compares the contents of two XML files and highlights any differences. It loads each document into a flattened list of parameter/value pairs, allowing you to quickly spot missing nodes, attribute changes, or mismatched values.

## Features

- **Side-by-side XML comparison** – Load a left and right XML file and run an asynchronous comparison to detect parameter differences. Missing values are highlighted in red inside the results grid.
- **Dynamic filtering** – Narrow the results with a free-text filter, limit the view to only parameters whose values differ, or ignore case when comparing values.
- **Copy parameter names** – Click any row to copy the parameter path to the clipboard and get an in-app toast notification confirming the action.
- **Excel export** – Export the currently visible differences to an `.xlsx` workbook so the report can be shared. The export uses the Open XML SDK.

## Prerequisites

- Windows 7 or later (required for WPF)
- [.NET 6 SDK](https://dotnet.microsoft.com/en-us/download/dotnet/6.0) with Windows desktop support
- (Optional) [Visual Studio 2022](https://visualstudio.microsoft.com/) with the ".NET desktop development" workload for an IDE experience

The project references the following NuGet packages:

- [`DocumentFormat.OpenXml`](https://www.nuget.org/packages/DocumentFormat.OpenXml/) for generating Excel exports.
- [`Notifications.Wpf`](https://www.nuget.org/packages/Notifications.Wpf/) for toast-style notifications.

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
2. Select **Compare** to generate the differences report.
3. Use the filter controls to refine the list:
   - Enter text in the **Filter** box to match by parameter path.
   - Enable **Only Show Different Parameter** to hide equal values.
   - Enable **Ignore Value Case** to treat values case-insensitively.
4. Review the grid for each parameter, its left value, and its right value. Rows with missing data are highlighted in red.
5. Click a row to copy its parameter path if you need to reference it elsewhere.
6. Press **Export** to save the filtered results to Excel.

The status text at the bottom summarizes the visible differences, including the total count and missing values on each side.

## How It Works

The comparer flattens each XML document by traversing every node, combining element names, indexes, and attributes into unique parameter paths. These paths and their values are compared to build the difference list displayed in the UI.

## License

This project is licensed under the [MIT License](LICENSE).
