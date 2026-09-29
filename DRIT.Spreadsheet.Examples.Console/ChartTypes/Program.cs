using System.IO.Compression;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Chart;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
var data = workbook.Worksheets[0];
data.Name = "Data";
data.GetRange("A1:B4").SetValue(new object[,] {
    { "Month", "Sales" },
    { 1, 120 },
    { 2, 150 },
    { 3, 175 }
});

var barSheet = workbook.AddWorksheet("Bar");
var bar = barSheet.Charts.Add<BarChart>("A1");
bar.DataSource = "Data!$A$1:$B$4";
var lineSheet = workbook.AddWorksheet("Line");
var line = lineSheet.Charts.Add<LineChart>("A1");
line.DataSource = "Data!$A$1:$B$4";
var pieSheet = workbook.AddWorksheet("Pie");
var pie = pieSheet.Charts.Add<PieChart>("A1");
pie.DataSource = "Data!$A$1:$B$4";
var scatterSheet = workbook.AddWorksheet("Scatter");
var scatter = scatterSheet.Charts.Add<ScatterChart>("A1");
scatter.AddSeries("Data!$B$1", "Data!$A$2:$A$4", "Data!$B$2:$B$4");

var path = ExampleSupport.OutputPath("ChartTypes.xlsx");
workbook.SaveAs(path);
ExampleSupport.Require(barSheet.Charts.ChartsByName.Count == 1, "The bar chart was not registered.");
ExampleSupport.Require(line.ChartType == ChartType.Line, "The line chart type was not configured.");
ExampleSupport.Require(scatter.DataSeries.Count == 1, "The scatter chart series was not configured.");
using var archive = ZipFile.OpenRead(path);
ExampleSupport.Require(archive.Entries.Count(entry => entry.FullName.StartsWith("xl/charts/chart")) >= 4, "The chart parts were not saved.");
Console.WriteLine("Bar, line, pie, and scatter charts saved.");