using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
worksheet.GetRange("A1:C6").SetValue(new object[,] {
    { "Department", "Owner", "Score" },
    { "Legal", "Fred", 85 },
    { "Marketing", "Alice", 92 },
    { "Finance", "Bob", 78 },
    { "Legal", "Carol", 88 },
    { "IT", "Dave", 75 }
});
var filter = worksheet.Filter("A1:C6");
filter.Values(0, false, "Legal");
filter.Apply();
var path = ExampleSupport.OutputPath("AutoFilter.xlsx");
workbook.SaveAs(path);

var loaded = new Workbook(path);
var result = loaded.Worksheets[0];
ExampleSupport.Require(result.Rows[1].IsHidden == false, "A matching filter row should remain visible.");
ExampleSupport.Require(result.Rows[3].IsHidden, "A non-matching filter row should be hidden.");
ExampleSupport.ReportRoundTrip(path, loaded);
Console.WriteLine("Filtered rows to the Legal department.");