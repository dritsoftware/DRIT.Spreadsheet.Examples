using System;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
worksheet["A1"].Value = "External site";
worksheet["A1"].Hyperlink = new Hyperlink
{
    ExternalUri = new Uri("https://example.com"),
    IsExternal = true,
    Tooltip = "Open example.com"
};
var second = workbook.AddWorksheet("Details");
worksheet["A2"].Value = "Internal details";
worksheet["A2"].Hyperlink = new Hyperlink { TargetCell = second["B2"], Tooltip = "Jump to details" };
var path = ExampleSupport.OutputPath("Hyperlinks.xlsx");
workbook.SaveAs(path);

var loaded = new Workbook(path);
var external = loaded.Worksheets[0]["A1"].Hyperlink;
ExampleSupport.Require(external != null && external.ExternalUri.ToString() == "https://example.com/", "The external hyperlink was not preserved.");
ExampleSupport.Require(loaded.Worksheets[0]["A2"].Hyperlink != null, "The internal hyperlink was not preserved.");
ExampleSupport.ReportRoundTrip(path, loaded);
Console.WriteLine($"External link: {external.ExternalUri}");