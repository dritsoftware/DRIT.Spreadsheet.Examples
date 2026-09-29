using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
worksheet.GetRange("A1:C4").SetValue(new object[,] {
    { "Product", "Region", "Units" },
    { "Atlas", "West", 4 },
    { "Meridian", "East", 7 },
    { "Orion", "North", 3 }
});
worksheet.Rows.Group(1, 3, false);
worksheet.Columns.Autofit();
worksheet.Rows.Autofit();
worksheet.Columns[0].WidthCharacters = 18;
worksheet.Columns[1].WidthCharacters = 14;
worksheet.Rows[0].HeightPoints = 24;

var path = ExampleSupport.OutputPath("RowColumnOperations.xlsx");
workbook.SaveAs(path);
var loaded = new Workbook(path);
var result = loaded.Worksheets[0];
ExampleSupport.Require(result.Columns[0].WidthCharacters == 18, "The product column width was not saved.");
ExampleSupport.Require(result.Rows[0].HeightPoints == 24, "The header row height was not saved.");
ExampleSupport.ReportRoundTrip(path, loaded);
Console.WriteLine($"Column A width: {result.Columns[0].WidthCharacters}");