using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
worksheet.Name = "Report";
worksheet["A1"].Value = "Quarterly report";
var headers = worksheet.PageSetup.HeaderFooters;
headers.DifferentOddEven = true;
headers.DifferentFirst = true;
headers.OddHeader.LeftSection = "DRIT Software";
headers.OddHeader.CenterSection = "Quarterly Report";
headers.OddFooter.RightSection = "Page &P of &N";
headers.EvenHeader.CenterSection = "Quarterly Report (Even)";
headers.FirstHeader.CenterSection = "Confidential";

var path = ExampleSupport.OutputPath("HeadersAndFooters.xlsx");
workbook.SaveAs(path);
var loaded = new Workbook(path);
var loadedHeaders = loaded.Worksheets[0].PageSetup.HeaderFooters;
ExampleSupport.Require(loadedHeaders.DifferentOddEven, "Odd/even headers were not preserved.");
ExampleSupport.Require(loadedHeaders.OddHeader.CenterSection == "Quarterly Report", "The odd header was not preserved.");
ExampleSupport.Require(loadedHeaders.OddFooter.RightSection == "Page &P of &N", "The footer fields were not preserved.");
ExampleSupport.ReportRoundTrip(path, loaded);
Console.WriteLine($"Header: {loadedHeaders.OddHeader.GetText()}");