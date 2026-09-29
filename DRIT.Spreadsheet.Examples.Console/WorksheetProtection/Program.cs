using System.IO.Compression;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
worksheet.Name = "Protected";
worksheet.GetRange("A1:B3").SetValue(new object[,] {
    { "Editable", "Locked" },
    { "User input", "Formula" },
    { "", "" }
});
worksheet["A2"].IsLocked = false;
worksheet.Protection.Protect("sheet-secret");
worksheet.Protection.InsertRows = false;
worksheet.Protection.DeleteColumns = false;
worksheet.Protection.UseAutoFilter = true;

var path = ExampleSupport.OutputPath("WorksheetProtection.xlsx");
workbook.SaveAs(path);
ExampleSupport.Require(worksheet.Protection.Protected, "The worksheet should be protected.");
ExampleSupport.Require(worksheet.Protection.PasswordHash != 0, "The worksheet password hash was not created.");
using var archive = ZipFile.OpenRead(path);
var sheetEntry = archive.Entries.First(entry => entry.FullName.StartsWith("xl/worksheets/sheet"));
using var sheetXml = new StreamReader(sheetEntry.Open());
ExampleSupport.Require(sheetXml.ReadToEnd().Contains("sheetProtection"), "Worksheet protection metadata was not saved.");
Console.WriteLine("Worksheet protection saved with A2 unlocked.");