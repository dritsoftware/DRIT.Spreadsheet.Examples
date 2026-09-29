using System.IO.Compression;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Chart;
using DRIT.Spreadsheet.Drawing;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
var data = workbook.Worksheets[0];
data.Name = "Data";
data.GetRange("A1:B4").SetValue(new object[,] {
    { "Region", "Sales" },
    { "East", 500 },
    { "West", 300 },
    { "North", 200 }
});
var chartsheet = workbook.AddChartsheet("SalesChart");
var chart = chartsheet.Charts.Add<BarChart>(1, "Sales Bar", new SizeEmu { Width = 5_029_200L, Height = 3_400_200L });
chart.DataSource = "Data!$A$1:$B$4";
chart.AddTitle();
chart.Title.FormattedText.Text = "Sales by region";

var path = ExampleSupport.OutputPath("ChartSheets.xlsx");
workbook.SaveAs(path);
ExampleSupport.Require(workbook.Sheets.Contains(chartsheet), "The chartsheet was not added to the workbook.");
ExampleSupport.Require(chartsheet.Charts.Count == 1, "The chartsheet chart was not registered.");
using var archive = ZipFile.OpenRead(path);
ExampleSupport.Require(archive.Entries.Any(entry => entry.FullName.Contains("chartsheet")), "The chartsheet part was not saved.");
Console.WriteLine("Dedicated sales chartsheet saved.");