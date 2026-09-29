using System.IO.Compression;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
var source = workbook.Worksheets[0];
source.Name = "Source";
source.GetRange("A1:B4").SetValue(new object[,] {
    { "Region", "Sales" },
    { "East", 500 },
    { "West", 300 },
    { "East", 200 }
});
var pivotSheet = workbook.AddWorksheet("Pivot");
var pivot = pivotSheet.PivotTables.Add("Source!A1:B4", "A1", "SalesByRegion");
pivot.DataFields.Add("Sales");
pivot.RowFields.Add("Region");
var slicer = pivotSheet.Slicers.Add(pivot, "D1", "Region");
var model = pivotSheet.Slicers[slicer];
model.Caption = "Region filter";
model.ColumnCount = 2;
model.StyleName = "SlicerStyleLight1";
var path = ExampleSupport.OutputPath("Slicers.xlsx");
workbook.SaveAs(path);

using var archive = ZipFile.OpenRead(path);
var slicerEntries = archive.Entries
    .Where(entry => entry.FullName.Contains("slicer", StringComparison.OrdinalIgnoreCase))
    .Select(entry => entry.FullName)
    .ToList();
ExampleSupport.Require(slicerEntries.Count > 0, "The slicer package parts were not saved.");
ExampleSupport.Require(model.Caption == "Region filter", "The slicer caption was not configured.");
ExampleSupport.Require(model.ColumnCount == 2, "The slicer column count was not configured.");
Console.WriteLine($"Slicer package parts: {slicerEntries.Count}");