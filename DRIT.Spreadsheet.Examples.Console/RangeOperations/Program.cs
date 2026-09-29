using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
worksheet.GetRange("A1:B2").SetValue(new object[,] { { "North", 10 }, { "South", 20 } });
worksheet["D1"].Value = worksheet["A1"].Value;
worksheet["E1"].Value = worksheet["B1"].Value;
worksheet["D2"].Value = worksheet["A2"].Value;
worksheet["E2"].Value = worksheet["B2"].Value;
worksheet.Merge("A4:E4");
worksheet["A4"].Value = "Merged summary";
worksheet.GetRange("A5:E5").Clear();

var path = ExampleSupport.OutputPath("RangeOperations.xlsx");
workbook.SaveAs(path);
var loaded = new Workbook(path);
var result = loaded.Worksheets[0];
ExampleSupport.Require((string)result["D1"].Value == "North", "The copied range was not saved.");
ExampleSupport.Require((double)result["E2"].Value == 20.0, "The copied numeric value was not saved.");
ExampleSupport.Require((string)result["A4"].Value == "Merged summary", "The merged range was not saved.");
ExampleSupport.ReportRoundTrip(path, loaded);
Console.WriteLine("Copied, merged, and cleared ranges.");