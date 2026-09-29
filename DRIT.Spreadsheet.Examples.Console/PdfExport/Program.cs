using DRIT.Pdf;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;
using DRIT.Spreadsheet.Export.Pdf;

var workbook = new Workbook();
var summary = workbook.Worksheets[0];
summary.Name = "Summary";
summary["A1"].Value = "Phase 4 PDF export";
summary["A2"].Value = 2026;
var details = workbook.AddWorksheet("Details");
details["A1"].Value = "Detail row";

var path = ExampleSupport.OutputPath("PdfExport.pdf");
workbook.SaveAsPdf(path);
ExampleSupport.Require(File.Exists(path), "The PDF file was not created.");
ExampleSupport.Require(new FileInfo(path).Length > 0, "The PDF file is empty.");
using var pdf = Document.Load(path);
ExampleSupport.Require(pdf.Pages.Count >= 2, "The PDF should contain both visible worksheets.");
Console.WriteLine($"PDF saved with {pdf.Pages.Count} pages.");
