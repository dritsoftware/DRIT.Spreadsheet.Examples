using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
worksheet.Name = "Report";
worksheet.GetRange("A1:D8").SetValue(new object[,] {
    { "Region", "Q1", "Q2", "Q3" },
    { "East", 10, 12, 14 },
    { "West", 8, 11, 13 },
    { "North", 6, 9, 12 },
    { "South", 7, 10, 11 },
    { "Total", 31, 42, 50 },
    { "", "", "", "" },
    { "Prepared", "2026", "", "" }
});
worksheet.PageSetup.Page.Orientation = Orientation.Landscape;
worksheet.PageSetup.Page.PaperSize = PaperSize.A4;
worksheet.PageSetup.Page.FitToWidth = 1;
worksheet.PageSetup.Page.FitToHeight = 0;
worksheet.PageSetup.Sheet.PrintGridlines = true;
worksheet.PageSetup.Sheet.PrintHeadings = true;

var path = ExampleSupport.OutputPath("PageSetup.xlsx");
workbook.SaveAs(path);
var loaded = new Workbook(path);
var loadedSheet = loaded.Worksheets[0];
ExampleSupport.Require(loadedSheet.PageSetup.Page.Orientation == Orientation.Landscape, "Landscape orientation was not preserved.");
ExampleSupport.Require(loadedSheet.PageSetup.Page.PaperSize == PaperSize.A4, "A4 paper size was not preserved.");
ExampleSupport.Require(loadedSheet.PageSetup.Sheet.PrintGridlines, "Print gridlines were not preserved.");
ExampleSupport.ReportRoundTrip(path, loaded);
Console.WriteLine("Landscape A4 layout with printed gridlines saved.");