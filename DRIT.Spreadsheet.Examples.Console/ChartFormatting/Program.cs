using System.IO.Compression;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Chart;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
worksheet.Name = "Chart";
worksheet.GetRange("A1:B4").SetValue(new object[,] {
    { "Month", "Revenue" },
    { "Jan", 120 },
    { "Feb", 150 },
    { "Mar", 175 }
});
var chart = worksheet.Charts.Add<BarChart>("D2");
chart.DataSource = "Chart!$A$1:$B$4";
chart.AddTitle();
chart.Title.FormattedText.Text = "Monthly revenue";
chart.AddLegend();
chart.Legend.LegendPosition = LegendPosition.Bottom;
chart.CategoryAxis.AddTitle();
chart.CategoryAxis.Title.FormattedText.Text = "Month";
chart.ValueAxis.AddTitle();
chart.ValueAxis.Title.FormattedText.Text = "Revenue";
chart.ValueAxis.MajorGridLines = new Gridlines();
chart.DataSeries[0].AddDataLabels();
chart.DataSeries[0].DataLabels.Position = DataLabelPosition.OutsideEnd;

var path = ExampleSupport.OutputPath("ChartFormatting.xlsx");
workbook.SaveAs(path);
ExampleSupport.Require(chart.Title.FormattedText.Text == "Monthly revenue", "The chart title was not configured.");
ExampleSupport.Require(chart.Legend.LegendPosition == LegendPosition.Bottom, "The chart legend position was not configured.");
ExampleSupport.Require(chart.ValueAxis.MajorGridLines != null, "The chart gridlines were not configured.");
using var archive = ZipFile.OpenRead(path);
ExampleSupport.Require(archive.Entries.Any(entry => entry.FullName.StartsWith("xl/charts/chart")), "The formatted chart part was not saved.");
Console.WriteLine("Chart title, legend, axes, gridlines, and labels saved.");