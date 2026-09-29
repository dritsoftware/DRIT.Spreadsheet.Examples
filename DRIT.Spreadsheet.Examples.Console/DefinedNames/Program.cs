using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
worksheet.Name = "Data";
worksheet.GetRange("A1:A3").SetValue(new[] { 10, 20, 30 });
workbook.DefinedNames.DefineName("TaxRate", "0.2");
workbook.DefinedNames.DefineName("SalesData", worksheet.GetRange("A1:A3"));
var path = ExampleSupport.OutputPath("DefinedNames.xlsx");
workbook.SaveAs(path);

var loaded = new Workbook(path);
ExampleSupport.Require(loaded.DefinedNames.TryGetWorkbookDefinedName("TaxRate") != null, "The formula defined name was not preserved.");
ExampleSupport.Require(loaded.DefinedNames.TryGetWorkbookDefinedName("SalesData") != null, "The range defined name was not preserved.");
ExampleSupport.ReportRoundTrip(path, loaded);
Console.WriteLine("Defined names: TaxRate, SalesData");