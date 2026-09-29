using System.IO.Compression;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
workbook.Worksheets[0]["A1"].Value = "Protected structure";
workbook.Protection = new WorkbookProtection();
workbook.Protection.Protect("structure-secret", lockStructure: true, lockWindows: false, lockRevision: false);

var path = ExampleSupport.OutputPath("WorkbookProtection.xlsx");
workbook.SaveAs(path);
ExampleSupport.Require(workbook.IsStructureLocked, "The workbook structure should be locked.");
ExampleSupport.Require(workbook.Protection.PasswordHash != 0, "The workbook protection password hash was not created.");
using var archive = ZipFile.OpenRead(path);
using var workbookXml = new StreamReader(archive.GetEntry("xl/workbook.xml").Open());
var xml = workbookXml.ReadToEnd();
ExampleSupport.Require(xml.Contains("workbookProtection"), "Workbook protection metadata was not saved.");
Console.WriteLine("Workbook structure protection saved.");