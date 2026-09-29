using System.IO;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
worksheet.GetRange("A1:B3").SetValue(new object[,] {
    { "Name", "Note" },
    { "Atlas", "Primary | owner" },
    { "Meridian", "Ready" }
});
var path = ExampleSupport.OutputPath("MarkdownExport.md");
workbook.SaveAs(path, FormatType.Markdown);
var markdown = File.ReadAllText(path);
ExampleSupport.Require(markdown.Contains("| Name | Note |"), "The Markdown header was not exported.");
ExampleSupport.Require(markdown.Contains("Primary \\| owner"), "Markdown pipe characters should be escaped.");
Console.WriteLine($"Markdown lines: {markdown.Split('\n').Length}");