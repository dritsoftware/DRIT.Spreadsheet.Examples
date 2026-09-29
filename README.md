# DRIT.Spreadsheet

A pure-managed, cross-platform .NET library for creating, reading, manipulating,
and converting spreadsheet documents: **XLSX**, **XLS**, **XLSB**, **CSV**, and
**TSV**. The library provides a formula engine, formatting, charts, tables,
pivot tables, protection, encryption, digital signatures, and export workflows
without requiring Microsoft Excel or another office installation.

[![.NET](https://img.shields.io/badge/.NET-net48%20%7C%20netstandard2.0%20%7C%20net6.0-blue.svg)](https://dotnet.microsoft.com/)
[![NuGet](https://img.shields.io/nuget/v/DRIT.Spreadsheet.svg)](https://www.nuget.org/packages/DRIT.Spreadsheet)
[![License: Commercial](https://img.shields.io/badge/license-Commercial-orange.svg)](#license)

## Public examples

This repository contains **58 runnable console examples** covering the normal
create, load, edit, save, reload, export, and protection workflows. Browse the
[complete example catalog](DRIT.Spreadsheet.Examples.Console/README.md) for
source code, generated-output expectations, package-based commands, and current
API limitations.

The catalog is organized into foundation, data and analysis, formatting and
layout, export and security, and API-gated workflows. Every example uses
deterministic data, writes generated files below `Out`, and verifies its result
through a reload, package inspection, or a concise invariant check.

---

## Table of Contents

- [What is DRIT.Spreadsheet?](#what-is-dritspreadsheet)
- [Features](#features)
- [Supported File Formats](#supported-file-formats)
- [Platform Independence](#platform-independence)
- [Get Started](#get-started)
- [Security](#security)
- [Public Console Examples](#public-console-examples)
- [Building and Validation](#building-and-validation)
- [Repository Contents](#repository-contents)
- [License](#license)

---

## What is DRIT.Spreadsheet?

DRIT.Spreadsheet is a pure-managed spreadsheet library implemented in C#. It
does not depend on Microsoft Excel, Office Interop, LibreOffice, or a
third-party spreadsheet engine. Its package model, Open XML reader and writer,
legacy binary workbook support, formula calculation, drawing model, and
export pipeline are maintained in the product source.

The library supports both ordinary spreadsheet automation and document-quality
workflows: calculated formulas, tables and pivots, charts and shapes, workbook
protection, AES-256 encryption, XAdES-BES signatures, PDF export, and image
rendering.

## Features

- **Workbook I/O:** create, load, save, and reload workbooks from files or
	streams, with password, warning, and cancellation options.
- **Cells and formulas:** typed values, merged cells, ranges, shared and array
	formulas, circular-reference detection, iterative calculation, and broad
	Excel-function coverage.
- **Data workflows:** find and replace, sorting, AutoFilter, data validation,
	DataTable import/export, JSON import/export, CSV/TSV, and Markdown export.
- **Formatting and layout:** number formats, fonts, borders, fills, rich text,
	conditional formatting, themes, AutoFit, row and column sizing, grouping,
	page setup, print settings, headers, footers, and page breaks.
- **Workbook structures:** worksheets, tables, pivot tables, slicers, defined
	names, hyperlinks, comments, threaded comments, custom XML, and OLE objects.
- **Charts and drawings:** chart sheets, bar/column, line, pie, scatter, and
	other chart types; shapes, pictures, form controls, anchoring, and sparklines.
- **Protection and security:** worksheet and workbook protection, write
	protection, ECMA-376 Agile AES-256 encryption, and XAdES-BES digital
	signatures.
- **Conversion and rendering:** PDF export, workbook and worksheet-range image
	export, HTML import/export, and export to GitHub Flavored Markdown.
- **Automation:** VBA project and module preservation and editing.

## Supported File Formats

| Format | Read | Write | Notes |
|---|:---:|:---:|---|
| **XLSX** | Yes | Yes | Office Open XML workbook |
| **XLS** | Yes | Preservation-only | Legacy Excel binary workbook |
| **XLSB** | Yes | Yes | Excel binary workbook |
| **CSV** | Yes | Yes | Comma-separated values |
| **TSV** | Yes | Yes | Tab-separated values |
| **PDF** | No | Yes | Via the integrated DRIT.Pdf companion |
| **PNG, BMP, JPEG, TIFF** | No | Yes | Workbook and range image export |
| **EMF** | No | Yes | Shared drawing export pipeline |
| **HTML** | Yes | Yes | Table-oriented import and export |
| **Markdown** | No | Yes | GitHub Flavored Markdown tables |
| **JSON** | Yes | Yes | Range and record import/export |

> PDF export, PDF signing, and image rendering use the integrated
> **DRIT.Pdf** and **DRIT.Drawing** companions.

## Platform Independence

DRIT.Spreadsheet targets **.NET Framework 4.8**, **.NET Standard 2.0**, and
**.NET 6.0**. It runs on Windows, Linux, and macOS where the target framework
and graphics dependencies are supported.

No Microsoft Excel installation or Office Interop assembly is required.

---

## Get Started

### Install from NuGet

```powershell
Install-Package DRIT.Spreadsheet
```

Or with the .NET CLI:

```bash
dotnet add package DRIT.Spreadsheet
```

### Create and save a workbook

```csharp
using DRIT.Spreadsheet;

var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];

worksheet.Cells["A1"].Value = "Product";
worksheet.Cells["B1"].Value = "Price";
worksheet.Cells["A2"].Value = "Atlas Server";
worksheet.Cells["B2"].Value = 4899.00;
worksheet.GetRange("A1:B1").Font.Bold = true;

workbook.SaveAs("output.xlsx");
```

### Load and inspect a workbook

```csharp
using DRIT.Spreadsheet;

var workbook = Workbook.Load("input.xlsx");
Console.WriteLine($"Worksheets: {workbook.Worksheets.Count}");

foreach (var cell in workbook.Worksheets[0].UsedRange.GetAllCells())
		Console.WriteLine($"[{cell.Row.Index},{cell.Column.Index}] = {cell.Value}");
```

### Create a chart

```csharp
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Charts;

var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
worksheet.GetRange("A1:B3").SetValue(new object[,]
{
		{ "Month", "Sales" },
		{ "Jan", 1000 },
		{ "Feb", 1500 }
});

var chart = worksheet.Charts.Add<BarChart>("D2");
chart.DataSource = "Sheet1!$A$1:$B$3";
chart.AddTitle();
chart.AddLegend();

workbook.SaveAs("chart.xlsx");
```

### Export to PDF or an image

```csharp
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Export.Image;

var workbook = Workbook.Load("input.xlsx");

workbook.SaveAsPdf("output.pdf");
workbook.SaveAsImage("output\\{sheet}_page{page}.png", new WorkbookImageSaveOptions
{
		Format = ImageExportFormat.Png,
		Dpi = 150,
		ForceGridLines = true
});
```

PDF and image export require the integrated renderer packages. PDF export also
requires a valid DRIT.Pdf license at runtime; a successful package restore does
not provide that license.

### Encrypt a workbook

```csharp
using DRIT.Spreadsheet;

var workbook = new Workbook();
workbook.Worksheets[0].Cells["A1"].Value = "Confidential";
workbook.SaveAs("encrypted.xlsx", new XlsxSaveOptions
{
		EncryptionMode = XlsxEncryptionMode.Ecma376Agile,
		Password = "open-password"
});

var loaded = Workbook.Load("encrypted.xlsx",
		new XlsxLoadOptions { Password = "open-password" });
```

---

## Security

| Capability | Implementation |
|---|---|
| AES-256 encryption | `XlsxSaveOptions` with `EncryptionMode.Ecma376Agile` and `Password` |
| Worksheet protection | `Worksheet.Protection.Protect()` with granular permission flags |
| Workbook protection | Structure and window protection with password hashing |
| Write protection | `ReadOnlyRecommended` and file-sharing metadata |
| XAdES-BES signing | `XlsxDigitalSignatureSaveOptions` with a PFX certificate and signer role |
| Validate signatures | `Workbook.ListDigitalSignatures(path)` |
| Remove signatures | `Workbook.RemoveDigitalSignatures(path)` |
| PDF signing and encryption | Options provided by the integrated `DRIT.Pdf` export workflow |

---

## Public Console Examples

The [DRIT.Spreadsheet examples catalog](DRIT.Spreadsheet.Examples.Console/README.md)
contains one focused project per workflow:

| Area | Representative examples |
|---|---|
| Foundation | `GettingStarted`, `Reading`, `Writing`, `Preservation`, `CellValues`, `RangeOperations`, `RowColumnOperations`, `WorksheetManagement` |
| Data and analysis | `Tables`, `AutoFilter`, `Sorting`, `DataValidation`, `FindAndReplace`, `Hyperlinks`, `DefinedNames`, `DataTableImportExport`, `CsvImportExport`, `PivotTables`, `Slicers`, `JsonImportExport`, `MarkdownExport` |
| Formatting and layout | `Formatting`, `StylesAndThemes`, `RichText`, `PageSetup`, `HeadersAndFooters`, `RightToLeftText`, `UnitConversion`, `ConditionalFormatting` |
| Charts and drawings | `Charts`, `ChartTypes`, `ChartFormatting`, `ChartSheets`, `Shapes`, `ShapeAnchoring`, `Pictures`, `FormControls`, `Sparklines`, `Ole` |
| Protection and metadata | `DocumentProperties`, `WorkbookProtection`, `WorksheetProtection`, `WriteProtection`, `Encryption`, `DigitalSignature`, `CustomXml`, `Vba` |
| Conversion and rendering | `HtmlImportExport`, `PdfExport`, `ImageExport`, `ProgressAndCancellation` |
| Other workbook features | `FreezeSplit`, `WorksheetView`, `Comments`, `ThreadedComments`, `Scenarios`, `EcmaFormulas` |

The Phase 5 examples `HtmlImportExport`, `ImageExport`, `ThreadedComments`, and
`Scenarios` use APIs that may require a newer package release and are marked
package-gated until the released package exposes the same surface. The catalog
also documents package-inspection checks for pivot tables, slicers,
charts, custom XML, signatures, and metadata where object-model reload is not
the complete verification boundary.

---

## Building and Validation

The examples use released NuGet packages and are intended to build from a
clean public checkout:

```powershell
dotnet build .\DRIT.Spreadsheet.Examples.Console\DRIT.Spreadsheet.Examples.sln -f net8.0 --nologo
```

The `DigitalSignature` and `Vba` examples use `net48` because their signing and
VBA paths are platform-gated. PDF and image examples require the integrated renderer
packages, and PDF execution requires a valid DRIT.Pdf license. Generated files
are written under `DRIT.Spreadsheet.Examples.Console/Out`, which is ignored by
Git.

---

## Repository Contents

| Path | Purpose |
|---|---|
| `DRIT.Spreadsheet.Examples.Console/` | Standalone projects consuming the NuGet packages |

## Documentation and Resources

- [NuGet package](https://www.nuget.org/packages/DRIT.Spreadsheet/)
- [Product page](https://www.dritsoftware.com/netspreadsheet)
- [API documentation](https://www.dritsoftware.com/docs/netspreadsheet/api/index.html)
- [Support forums](https://www.dritsoftware.com/forums/)

## License

DRIT.Spreadsheet is a **closed-source, commercial product** of DR-IT Ltd. This
repository contains examples and documentation only. The examples are provided
for evaluation and learning. See the
[product page](https://www.dritsoftware.com/netspreadsheet) for licensing terms.

[Home](https://www.dritsoftware.com) | [Product Page](https://www.dritsoftware.com/netspreadsheet) | [NuGet](https://www.nuget.org/packages/DRIT.Spreadsheet)
