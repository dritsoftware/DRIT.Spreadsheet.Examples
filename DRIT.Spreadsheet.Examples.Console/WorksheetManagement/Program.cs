using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
var source = workbook.Worksheets[0];
source.Name = "Source";
source["A1"].Value = "Copied worksheet";
var copied = source.CopyTo("Copy");
var hidden = workbook.AddWorksheet("Hidden");
hidden.Visibility = SheetState.Hidden;
var path = ExampleSupport.OutputPath("WorksheetManagement.xlsx");
workbook.SaveAs(path);

var loaded = new Workbook(path);
ExampleSupport.Require(loaded.Worksheets.Count == 3, "The workbook should contain three worksheets.");
ExampleSupport.Require((string)loaded.Worksheets["Copy"]["A1"].Value == "Copied worksheet", "The copied worksheet was not saved.");
ExampleSupport.Require(loaded.Worksheets["Hidden"].Visibility == SheetState.Hidden, "The hidden worksheet state was not saved.");
ExampleSupport.ReportRoundTrip(path, loaded);
Console.WriteLine($"Created: {source.Name}, {copied.Name}, {hidden.Name}");