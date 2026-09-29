using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Drawing;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
var paragraph = new RichTextParagraph();
paragraph.Runs.Add(new RichTextRun("Quarterly "));
paragraph.Runs.Add(new RichTextRun("revenue") { Bold = true, Foreground = SpreadsheetColor.Accent1 });
paragraph.Runs.Add(new RichTextRun(" increased") { Italic = true, FontSize = 16 });
worksheet["A1"].RichText = paragraph;

var path = ExampleSupport.OutputPath("RichText.xlsx");
workbook.SaveAs(path);
var loaded = new Workbook(path);
var richText = loaded.Worksheets[0]["A1"].RichText;
ExampleSupport.Require(richText != null, "Rich text was not loaded.");
ExampleSupport.Require(richText.Runs.Count == 3, "Rich text runs were not preserved.");
ExampleSupport.Require(richText.Runs[1].Bold, "The bold rich text run was not preserved.");
ExampleSupport.Require(richText.Text == "Quarterly revenue increased", "Rich text content was not preserved.");
ExampleSupport.ReportRoundTrip(path, loaded);
Console.WriteLine($"Rich text runs: {richText.Runs.Count}");