using System.Linq;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

var csvPath = ExampleSupport.OutputPath("CsvImportExport.csv");
var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
worksheet.GetRange("A1:C3").SetValue(new object[,] {
    { "Name", "Age", "Note" },
    { "John", 30, "Engineer" },
    { "Jane", 35, "Smith, Jane" }
});
workbook.SaveAs(csvPath, FormatType.Csv);

var loaded = new Workbook(csvPath);
var result = loaded.Worksheets[0];
var lines = System.IO.File.ReadAllLines(csvPath);
ExampleSupport.Require(lines.Length >= 3, "The CSV should contain a header and two data rows.");
ExampleSupport.Require((string)result["C3"].Value == "Smith, Jane", "Quoted CSV text was not reloaded correctly.");
ExampleSupport.Require(lines.Any(line => line.Contains("Jane")), "The CSV output should contain Jane.");
ExampleSupport.ReportRoundTrip(csvPath, loaded);
Console.WriteLine($"CSV rows: {lines.Length - 1}");