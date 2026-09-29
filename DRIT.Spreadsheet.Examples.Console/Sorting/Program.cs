using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
worksheet.GetRange("A1:B4").SetValue(new object[,] {
    { "Department", "Owner" },
    { "B", "Zara" },
    { "A", "Mike" },
    { "A", "Alex" }
});
var sort = new SortState(worksheet.GetRange("A2:B4"));
sort.SortConditions.Add(new SortCondition(sort) { Index = 0, Descending = false });
sort.SortConditions.Add(new SortCondition(sort) { Index = 1, Descending = false });
sort.Sort();
var path = ExampleSupport.OutputPath("Sorting.xlsx");
workbook.SaveAs(path);

var loaded = new Workbook(path);
var result = loaded.Worksheets[0];
ExampleSupport.Require((string)result["A2"].Value == "A", "The primary sort key was not applied.");
ExampleSupport.Require((string)result["B2"].Value == "Alex", "The secondary sort key was not applied.");
ExampleSupport.ReportRoundTrip(path, loaded);
Console.WriteLine($"First sorted row: {result["A2"].Value}, {result["B2"].Value}");