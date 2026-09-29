using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
worksheet.Name = "RTL";
worksheet.View.RightToLeft = true;
worksheet["A1"].Value = "مرحبا بالعالم";
worksheet["A1"].Alignment.TextDirection = TextDirection.RightToLeft;
worksheet["B1"].Value = "LTR reference";
worksheet["B1"].Alignment.TextDirection = TextDirection.LeftToRight;

var path = ExampleSupport.OutputPath("RightToLeftText.xlsx");
workbook.SaveAs(path);
var loaded = new Workbook(path);
var loadedSheet = loaded.Worksheets[0];
ExampleSupport.Require(loadedSheet.View.RightToLeft, "Right-to-left sheet view was not preserved.");
ExampleSupport.Require(loadedSheet["A1"].Alignment.TextDirection == TextDirection.RightToLeft, "Right-to-left cell text was not preserved.");
ExampleSupport.ReportRoundTrip(path, loaded);
Console.WriteLine("RTL sheet and cell text saved.");