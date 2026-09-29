using System.IO;
using System.Text;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
worksheet.Name = "Report";
worksheet["A1"].Value = "Name";
worksheet["B1"].Value = "Score";
worksheet["A2"].Value = "Alice";
worksheet["B2"].Value = 95;
worksheet["A3"].Value = "Bob";
worksheet["B3"].Value = 87;

var exportPath = ExampleSupport.OutputPath("HtmlImportExport.html");
workbook.SaveAs(exportPath, new HtmlSaveOptions
{
    ExportGridLines = true,
    ExportHeadings = true
});
var html = File.ReadAllText(exportPath);
ExampleSupport.Require(html.Contains("<!DOCTYPE html>"), "The HTML doctype was not written.");
ExampleSupport.Require(html.Contains("Alice") && html.Contains("Bob"), "The HTML table values were not written.");

var importPath = ExampleSupport.OutputPath("HtmlImportSource.html");
File.WriteAllText(importPath,
    "<html><body><table><tr><td>Name</td><td>Score</td></tr>" +
    "<tr><td>Carol</td><td>91</td></tr></table></body></html>",
    new UTF8Encoding(false));
var imported = Workbook.Load(importPath);
ExampleSupport.Require(imported.Worksheets[0]["A1"].Value?.ToString() == "Name", "The HTML header was not imported.");
ExampleSupport.Require(imported.Worksheets[0]["A2"].Value?.ToString() == "Carol", "The HTML text value was not imported.");
ExampleSupport.Require(imported.Worksheets[0]["B2"].Value?.ToString() == "91", "The HTML numeric value was not imported.");
Console.WriteLine("HTML export and import completed with table values preserved.");
