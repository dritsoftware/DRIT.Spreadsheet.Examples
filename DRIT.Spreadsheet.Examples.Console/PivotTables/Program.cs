using System.IO.Compression;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;
using DRIT.Spreadsheet.Pivot;

var workbook = new Workbook();
var source = workbook.Worksheets[0];
source.Name = "Source";
source.GetRange("A1:B5").SetValue(new object[,] {
    { "Region", "Sales" },
    { "East", 500 },
    { "West", 300 },
    { "East", 200 },
    { "West", 100 }
});
var pivotSheet = workbook.AddWorksheet("Pivot");
var pivot = pivotSheet.PivotTables.Add("Source!A1:B5", "A1", "SalesByRegion");
var sales = pivot.DataFields.Add("Sales");
sales.AggregationFunction = AggregationFunction.Sum;
pivot.RowFields.Add("Region");
pivot.ShowGrandTotalsForRows = true;
pivot.Calculate();
var path = ExampleSupport.OutputPath("PivotTables.xlsx");
workbook.SaveAs(path);

var entries = ZipFile.OpenRead(path).Entries.Select(entry => entry.FullName).ToList();
ExampleSupport.Require(entries.Any(entry => entry.Contains("pivotCacheDefinition")), "The pivot cache definition part was not saved.");
ExampleSupport.Require(entries.Any(entry => entry.Contains("pivotTable")), "The pivot table definition part was not saved.");
ExampleSupport.Require(pivotSheet["A2"].Value != null, "The calculated pivot output should contain a row label.");
Console.WriteLine($"Pivot parts saved: {entries.Count(entry => entry.Contains("pivot"))}");