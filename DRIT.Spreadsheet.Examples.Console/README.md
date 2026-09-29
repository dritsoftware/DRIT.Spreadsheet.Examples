# DRIT.Spreadsheet examples

These focused console projects demonstrate the public DRIT.Spreadsheet API for
creating, reading, editing, and saving spreadsheet files. Each example writes
generated files to the repository-level `Out` directory and does not require an
Excel installation or vendor fixture.

## Requirements

- .NET Framework 4.8 or .NET 8 SDK.
- The examples use the released `DRIT.Spreadsheet` NuGet package.
- PDF and image examples also require the compatible `DRIT.Pdf` and
	`DRIT.Drawing` packages.

The examples target `net48` and `net8.0`. Use the framework supported by your
application and the package assets available in your environment.

## Run

From this directory:

```powershell
dotnet build .\GettingStarted\GettingStarted.csproj -f net8.0 --nologo
dotnet run --project .\GettingStarted\GettingStarted.csproj -f net8.0
```

Package-mode execution requires the normal DRIT.Spreadsheet license. PDF
execution additionally requires a valid DRIT.Pdf license; NuGet authentication
does not grant a product license.

Build the complete example solution with:

```powershell
dotnet build .\DRIT.Spreadsheet.Examples.sln -f net8.0 --nologo
```

The package references use the approved rolling 26.x policy. Record the exact
resolved package versions when validating a release.

## Examples

### Foundation

| Example | Demonstrates |
| --- | --- |
| [GettingStarted](GettingStarted) | Create, format, save, reload, and inspect a workbook |
| [Reading](Reading) | Generate a workbook, load it, and read typed values |
| [Writing](Writing) | Write scalar and array data and merge cells |
| [Preservation](Preservation) | Edit a reloaded workbook while preserving existing content |
| [CellValues](CellValues) | Text, numbers, booleans, dates, and formulas |
| [RangeOperations](RangeOperations) | Bulk values, copy, clear, and merge ranges |
| [RowColumnOperations](RowColumnOperations) | Width, height, autofit, and grouping |
| [WorksheetManagement](WorksheetManagement) | Rename, copy, add, and hide worksheets |

### Existing API examples

| Example | Demonstrates |
| --- | --- |
| [Charts](Charts) | Chart creation and chart types |
| [Comments](Comments) | Cell comments |
| [ConditionalFormatting](ConditionalFormatting) | Conditional formatting rules |
| [Filter](Filter) | Worksheet filtering |
| [FormControls](FormControls) | Form controls |
| [Formatting](Formatting) | Cell and range formatting |
| [EcmaFormulas](EcmaFormulas) | Formula evaluation |
| [FreezeSplit](FreezeSplit) | Freeze panes and split views |
| [Ole](Ole) | Embedded OLE objects |
| [Pictures](Pictures) | Pictures |
| [Shapes](Shapes) | Drawing shapes |
| [Sparklines](Sparklines) | Sparklines |
| [Vba](Vba) | VBA project preservation |
| [WorksheetView](WorksheetView) | Worksheet view settings |

### Phase 2 workflows

| Example | Demonstrates |
| --- | --- |
| [Tables](Tables) | Create an Excel table and calculated column |
| [AutoFilter](AutoFilter) | Apply a value filter and verify filtered rows |
| [Sorting](Sorting) | Multi-key ascending and descending sort conditions |
| [DataValidation](DataValidation) | List and decimal validation rules |
| [FindAndReplace](FindAndReplace) | Workbook find and replace |
| [Hyperlinks](Hyperlinks) | External and internal hyperlinks |
| [DefinedNames](DefinedNames) | Formula and range defined names |
| [DataTableImportExport](DataTableImportExport) | Import and export `DataTable` values |
| [CsvImportExport](CsvImportExport) | CSV save, reload, quoting, and values |
| [PivotTables](PivotTables) | Pivot fields, calculation, and saved pivot package parts |
| [Slicers](Slicers) | Pivot slicer configuration and saved slicer package parts |
| [JsonImportExport](JsonImportExport) | JSON record import and range export |
| [MarkdownExport](MarkdownExport) | Markdown table export with escaped pipes |

The Phase 2 examples use deterministic generated data and write their results
to `Out`. Pivot and slicer examples inspect the saved XLSX package because
these features are represented by related Open XML parts. The current source
loader does not hydrate saved slicers back into the worksheet slicer collection,
so the slicer example verifies configuration before save and the persisted
package parts after save.

### Phase 3 formatting and layout

| Example | Demonstrates |
| --- | --- |
| [StylesAndThemes](StylesAndThemes) | Reusable cell styles, theme colors, and a custom workbook theme |
| [RichText](RichText) | Multiple formatted runs in one cell and reload |
| [PageSetup](PageSetup) | Landscape A4 layout, print scaling, headings, and gridlines |
| [HeadersAndFooters](HeadersAndFooters) | Odd/even/first-page sections and page fields |
| [RightToLeftText](RightToLeftText) | RTL worksheet view and cell text direction |
| [UnitConversion](UnitConversion) | Point, pixel, centimeter, inch, and row-height conversion |
| [ChartTypes](ChartTypes) | Bar, line, pie, and scatter charts |
| [ChartFormatting](ChartFormatting) | Chart titles, legends, axes, gridlines, and data labels |
| [ChartSheets](ChartSheets) | Dedicated chartsheets with full-sheet charts |
| [ShapeAnchoring](ShapeAnchoring) | Cell-relative shape placement with pixel offsets and sizes |

Chart and chartsheet examples inspect saved XLSX chart parts in addition to
checking the public chart model. The current print-area and print-title setters
are not presented as persisted features because generated print defined names
are not emitted by the current writer.

### Phase 4 export, metadata, and security

| Example | Demonstrates |
| --- | --- |
| [DocumentProperties](DocumentProperties) | Summary and custom document properties with package inspection |
| [WorkbookProtection](WorkbookProtection) | Workbook structure protection and password hashing |
| [WorksheetProtection](WorksheetProtection) | Worksheet protection flags and unlocked cells |
| [WriteProtection](WriteProtection) | Read-only recommendation metadata |
| [Encryption](Encryption) | ECMA-376 agile encryption, correct-password load, and wrong-password rejection |
| [DigitalSignature](DigitalSignature) | Create, list, and remove an XLSX digital signature |
| [CustomXml](CustomXml) | Add, select, edit, and persist a custom XML part |
| [PdfExport](PdfExport) | Export multiple worksheets to PDF and verify page count |
| [ProgressAndCancellation](ProgressAndCancellation) | PDF page progress events and pre-cancelled export behavior |

Phase 4 examples use `net8.0`, except `DigitalSignature` and `Vba`, which run on
`net48` because their signing and VBA paths are platform-gated in the current
stack. PDF examples require a valid DRIT.Pdf license. The document-properties writer currently persists the author and custom
property package parts, but title/subject/keywords/company and custom-property
type hydration are not asserted as XLSX reload behavior. Workbook protection
currently emits the protection element while the structure-lock state is
verified through the public in-memory model.

### Phase 5 API-gated workflows

| Example | Demonstrates | Package status |
| --- | --- | --- |
| [HtmlImportExport](HtmlImportExport) | HTML table export and HTML table import | Local source only; pending the released package API |
| [ImageExport](ImageExport) | PNG page rendering with dimensions and file checks | Local source only; requires current DRIT.Drawing/image-export APIs |
| [ThreadedComments](ThreadedComments) | Authors, threaded replies, resolved comments, and XLSX reload | Local source only; pending the released package API |
| [Scenarios](Scenarios) | Scenario metadata, input cells, and XLSX reload | Local source only; pending the released package loader API |

The Phase 5 examples use APIs that may require a newer package release than the
currently resolved 26.x version. Build those projects individually and treat
unsupported APIs as package-compatibility issues rather than source-checkout
requirements. Query tables, streaming load/save, fixed-width text, and other
backlog items remain planned because there is no stable public authoring path
for a deterministic console example.

## Files and output

`In` contains small checked-in input fixtures when an example needs one. Every
generated workbook is written below `Out`; the directory is safe to delete and
recreate.

## Resources

- [NuGet package](https://www.nuget.org/packages/DRIT.Spreadsheet/)
- [Product page](https://www.dritsoftware.com/netspreadsheet)
- [API documentation](https://www.dritsoftware.com/docs/netspreadsheet/api/index.html)
- [Support forums](https://www.dritsoftware.com/forums/)
