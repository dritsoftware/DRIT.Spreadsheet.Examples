using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;
using System.IO.Compression;

var workbook = new Workbook();
workbook.DocumentProperties.Summary.Title = "Quarterly Report";
workbook.DocumentProperties.Summary.Author = "DRIT Examples";
workbook.DocumentProperties.Summary.Subject = "Spreadsheet metadata";
workbook.DocumentProperties.Summary.Keywords = "spreadsheet, metadata, phase4";
workbook.DocumentProperties.Summary.Company = "DRIT Software";
workbook.DocumentProperties.Custom.Add(new DocumentProperty
{
    Name = "ReviewStatus",
    Value = "Approved"
});

var path = ExampleSupport.OutputPath("DocumentProperties.xlsx");
workbook.SaveAs(path);
using var archive = ZipFile.OpenRead(path);
using var corePropertiesXml = new StreamReader(archive.GetEntry("docProps/core.xml").Open());
var coreXml = corePropertiesXml.ReadToEnd();
ExampleSupport.Require(coreXml.Contains("DRIT Examples"), "The document author was not saved.");
using var customPropertiesXml = new StreamReader(archive.GetEntry("docProps/custom.xml").Open());
var customXml = customPropertiesXml.ReadToEnd();
ExampleSupport.Require(customXml.Contains("ReviewStatus"), "The custom property name was not saved.");
ExampleSupport.Require(customXml.Contains("Approved"), "The custom property value was not saved.");
Console.WriteLine("Summary and custom document properties saved.");