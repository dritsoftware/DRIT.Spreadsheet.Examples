using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
worksheet.Name = "Sales";
worksheet["A1"].Value = "Product";
worksheet["B1"].Value = "Units";
worksheet["C1"].Value = "Revenue";
worksheet.GetRange("A1:C1").Font.Bold = true;
worksheet["A2"].Value = "Atlas Server";
worksheet["B2"].Value = 4;
worksheet["C2"].Value = 4899.0;
worksheet["A3"].Value = "Meridian Cloud";
worksheet["B3"].Value = 7;
worksheet["C3"].Value = 3299.0;
worksheet["C4"].Formula = "=SUM(C2:C3)";

var path = ExampleSupport.OutputPath("GettingStarted.xlsx");
workbook.SaveAs(path);
var loaded = new Workbook(path);

ExampleSupport.Require(loaded.Sheets.Count == 1, "The workbook should contain one worksheet.");
ExampleSupport.Require(Convert.ToString(loaded.Worksheets[0]["A2"].Value) == "Atlas Server", "The saved product was not reloaded.");
ExampleSupport.Require(loaded.Worksheets[0]["C4"].Formula == "=SUM(C2:C3)", "The saved formula was not reloaded.");
ExampleSupport.ReportRoundTrip(path, loaded);
Console.WriteLine($"First product: {Convert.ToString(loaded.Worksheets[0]["A2"].Value)}");