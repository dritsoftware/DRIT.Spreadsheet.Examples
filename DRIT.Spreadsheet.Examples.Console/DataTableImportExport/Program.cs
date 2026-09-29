using System;
using System.Data;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

var input = new DataTable();
input.Columns.Add("ID", typeof(int));
input.Columns.Add("Name", typeof(string));
input.Rows.Add(100, "John");
input.Rows.Add(101, "Jane");
var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
DataImportExport.Import(worksheet["A1"], input, new DataTableImportOptions { ImportColumnNames = true });
var exported = DataImportExport.ExportRangeToDataTable(worksheet.GetRange("A1:B3"));
var path = ExampleSupport.OutputPath("DataTableImportExport.xlsx");
workbook.SaveAs(path);

var loaded = new Workbook(path);
var reloaded = DataImportExport.ExportRangeToDataTable(loaded.Worksheets[0].GetRange("A1:B3"));
ExampleSupport.Require(reloaded.Rows.Count == 2, "The imported data rows were not preserved.");
ExampleSupport.Require(Convert.ToInt32(reloaded.Rows[1]["ID"]) == 101, "The exported DataTable value was not preserved.");
ExampleSupport.Require(exported.Columns.Count == 2, "The in-memory DataTable export should have two columns.");
ExampleSupport.ReportRoundTrip(path, loaded);
Console.WriteLine($"DataTable rows: {reloaded.Rows.Count}");