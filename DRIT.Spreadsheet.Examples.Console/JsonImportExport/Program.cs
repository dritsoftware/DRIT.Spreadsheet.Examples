using System.Text.Json;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

var json = "[{\"Name\":\"Atlas\",\"Units\":4},{\"Name\":\"Meridian\",\"Units\":7}]";
var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
DataImportExport.Import(worksheet["A1"], json);
var exported = DataImportExport.ExportRangeToJson(worksheet.GetRange("A1:B3"));
var path = ExampleSupport.OutputPath("JsonImportExport.xlsx");
workbook.SaveAs(path);

using var document = JsonDocument.Parse(exported);
ExampleSupport.Require(document.RootElement.GetArrayLength() == 2, "JSON export should contain two records.");
ExampleSupport.Require(document.RootElement[1].GetProperty("Name").GetString() == "Meridian", "The second JSON record was not exported.");
var loaded = new Workbook(path);
ExampleSupport.Require((string)loaded.Worksheets[0]["A2"].Value == "Atlas", "The imported JSON value was not saved.");
ExampleSupport.ReportRoundTrip(path, loaded);
Console.WriteLine($"JSON records: {document.RootElement.GetArrayLength()}");