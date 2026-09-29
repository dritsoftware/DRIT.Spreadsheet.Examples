using System.IO.Compression;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
workbook.Worksheets[0]["A1"].Value = "Read-only recommendation";
workbook.ReadOnlyRecommended = true;
var path = ExampleSupport.OutputPath("WriteProtection.xlsx");
workbook.SaveAs(path);

ExampleSupport.Require(workbook.ReadOnlyRecommended, "Read-only recommendation was not configured.");
using var archive = ZipFile.OpenRead(path);
using var workbookXml = new StreamReader(archive.GetEntry("xl/workbook.xml").Open());
var xml = workbookXml.ReadToEnd();
ExampleSupport.Require(xml.Contains("fileSharing"), "The file-sharing metadata was not saved.");
ExampleSupport.Require(xml.Contains("readOnlyRecommended=\"1\""), "The read-only recommendation was not saved.");
Console.WriteLine("Read-only recommendation saved.");