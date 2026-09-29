using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
worksheet["A1"].Value = "Text";
worksheet["A2"].Value = 42;
worksheet["A3"].Value = 12.5;
worksheet["A4"].Value = true;
worksheet["A5"].Value = new DateTime(2026, 9, 28);
worksheet["B1"].Formula = "=SUM(A2:A3)";
workbook.Calculate();

var path = ExampleSupport.OutputPath("CellValues.xlsx");
workbook.SaveAs(path);
var loaded = new Workbook(path);
var result = loaded.Worksheets[0];
ExampleSupport.Require(result["A1"].ValueType == DRIT.Spreadsheet.ValueType.String, "Text should retain its value type.");
ExampleSupport.Require(result["A2"].ValueType == DRIT.Spreadsheet.ValueType.Number, "Numbers should retain their value type.");
ExampleSupport.Require(result["A4"].ValueType == DRIT.Spreadsheet.ValueType.Boolean, "Booleans should retain their value type.");
ExampleSupport.Require(result["B1"].Formula == "=SUM(A2:A3)", "The formula should be preserved.");
ExampleSupport.ReportRoundTrip(path, loaded);
Console.WriteLine($"Types: {result["A1"].ValueType}, {result["A2"].ValueType}, {result["A4"].ValueType}");