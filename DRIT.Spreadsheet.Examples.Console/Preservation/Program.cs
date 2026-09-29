using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
worksheet.Name = "Preserved";
worksheet["A1"].Value = "Keep this value";
worksheet["A1"].Font.Bold = true;
worksheet["B1"].Value = 2026;
var path = ExampleSupport.OutputPath("Preservation.xlsx");
workbook.SaveAs(path);

var edited = new Workbook(path);
edited.Worksheets[0]["C1"].Value = "Added after reload";
edited.SaveAs(path);
var loaded = new Workbook(path);

ExampleSupport.Require((string)loaded.Worksheets[0]["A1"].Value == "Keep this value", "Existing text was not preserved.");
ExampleSupport.Require(loaded.Worksheets[0]["A1"].Font.Bold, "Existing formatting was not preserved.");
ExampleSupport.Require((double)loaded.Worksheets[0]["B1"].Value == 2026.0, "Existing numeric data was not preserved.");
ExampleSupport.Require((string)loaded.Worksheets[0]["C1"].Value == "Added after reload", "The edit was not saved.");
ExampleSupport.ReportRoundTrip(path, loaded);
Console.WriteLine("Original content and a post-reload edit survived two saves.");