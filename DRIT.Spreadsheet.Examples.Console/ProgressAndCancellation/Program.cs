using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;
using DRIT.Spreadsheet.Export;
using DRIT.Spreadsheet.Export.Pdf;

var workbook = new Workbook();
workbook.Worksheets[0]["A1"].Value = "Progress-enabled export";
var reports = new List<SaveProgress>();
var outputPath = ExampleSupport.OutputPath("ProgressAndCancellation.pdf");
workbook.SaveAsPdf(outputPath, new WorkbookPdfSaveOptions
{
    Progress = new Progress<SaveProgress>(reports.Add)
});
ExampleSupport.Require(reports.Any(report => report.Stage == SaveProgressStage.PageSaving), "PDF page-start progress was not reported.");
ExampleSupport.Require(reports.Any(report => report.Stage == SaveProgressStage.PageSaved), "PDF page-complete progress was not reported.");

using var cancelledStream = new MemoryStream();
using var cancellation = new CancellationTokenSource();
cancellation.Cancel();
try
{
    workbook.SaveAsPdf(cancelledStream, cancellation.Token);
    throw new InvalidOperationException("A pre-cancelled PDF export should be rejected.");
}
catch (OperationCanceledException)
{
    Console.WriteLine($"PDF export reported {reports.Count} progress events and honored cancellation.");
}
