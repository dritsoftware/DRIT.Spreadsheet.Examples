using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

var source = new Workbook();
source.Worksheets[0].Name = "Orders";
source.AddWorksheet("Summary");
source.Worksheets[0]["A1"].Value = "Order";
source.Worksheets[0]["B1"].Value = "Amount";
source.Worksheets[0]["A2"].Value = "SO-1001";
source.Worksheets[0]["B2"].Value = 1250.50;
var path = ExampleSupport.OutputPath("Reading.xlsx");
source.SaveAs(path);

var workbook = new Workbook(path);
ExampleSupport.Require(workbook.Worksheets.Count == 2, "The saved workbook should contain two worksheets.");
ExampleSupport.Require((string)workbook.Worksheets[0]["A2"].Value == "SO-1001", "The order value was not read.");
ExampleSupport.Require((double)workbook.Worksheets[0]["B2"].Value == 1250.50, "The amount was not read.");
ExampleSupport.ReportRoundTrip(path, workbook);
Console.WriteLine($"Read {workbook.Worksheets[0]["A2"].Value}: {workbook.Worksheets[0]["B2"].Value}");