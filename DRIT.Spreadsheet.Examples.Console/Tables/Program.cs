using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
worksheet.Name = "Sales";
worksheet.GetRange("A1:C4").SetValue(new object[,] {
    { "Product", "Price", "Quantity" },
    { "Apple", 1.5, 4 },
    { "Banana", 0.9, 7 },
    { "Cherry", 2.25, 3 }
});
var table = worksheet.Tables.Add("A1:C4", "SalesTable", true);
table.Columns[2].CalculatedFormula = "=[Price]*[Quantity]";
ExampleSupport.Require(table.Columns[2].CalculatedFormula == "=[Price]*[Quantity]", "The calculated column formula was not registered.");
var path = ExampleSupport.OutputPath("Tables.xlsx");
workbook.SaveAs(path);

var loaded = new Workbook(path);
var savedTable = loaded.Worksheets[0].Tables[0];
ExampleSupport.Require(savedTable.Name == "SalesTable", "The table name was not preserved.");
ExampleSupport.Require(savedTable.Columns.Count == 3, "The table columns were not preserved.");
ExampleSupport.ReportRoundTrip(path, loaded);
Console.WriteLine($"Table: {savedTable.Name}, columns: {savedTable.Columns.Count}");