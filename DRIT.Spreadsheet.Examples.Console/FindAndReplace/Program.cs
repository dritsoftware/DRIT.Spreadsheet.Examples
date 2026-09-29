using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;
using DRIT.Spreadsheet.FindReplace;

var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
worksheet["A1"].Value = "Apollo 1";
worksheet["A2"].Value = "Apollo 2";
worksheet["A3"].Value = "Luna";
var matches = workbook.FindAll(new FindArguments { FindWhat = "Apollo" }).ToList();
workbook.ReplaceAll(new ReplaceArguments { FindWhat = "Apollo", ReplaceWith = "Orion" });
var path = ExampleSupport.OutputPath("FindAndReplace.xlsx");
workbook.SaveAs(path);

var loaded = new Workbook(path);
var result = loaded.Worksheets[0];
ExampleSupport.Require(matches.Count == 2, "The find operation should return two matches.");
ExampleSupport.Require((string)result["A1"].Value == "Orion 1", "The first replacement was not preserved.");
ExampleSupport.Require((string)result["A3"].Value == "Luna", "An unrelated value should not be replaced.");
ExampleSupport.ReportRoundTrip(path, loaded);
Console.WriteLine($"Matches replaced: {matches.Count}");