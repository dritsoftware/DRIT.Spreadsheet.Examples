using System.IO;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

const string password = "phase4-secret";
var workbook = new Workbook();
workbook.Worksheets[0][0, 0].Value = "Encrypted workbook";
var path = ExampleSupport.OutputPath("Encryption.xlsx");
workbook.SaveAs(path, new XlsxSaveOptions
{
    EncryptionMode = XlsxEncryptionMode.Ecma376Agile,
    Password = password
});

var loaded = Workbook.Load(path, new XlsxLoadOptions
{
    EncryptionExpectation = XlsxEncryptionExpectation.Ecma376Agile,
    Password = password
});
var reloadedValue = loaded.Worksheets[0][0, 0].Value?.ToString();
ExampleSupport.Require(string.Equals(reloadedValue, "Encrypted workbook", StringComparison.Ordinal), "The encrypted workbook did not load with the correct password.");

try
{
    Workbook.Load(path, new XlsxLoadOptions
    {
        EncryptionExpectation = XlsxEncryptionExpectation.Ecma376Agile,
        Password = "wrong-password"
    });
    throw new InvalidOperationException("An incorrect password should be rejected.");
}
catch (InvalidDataException)
{
    Console.WriteLine("Encrypted workbook loaded with the correct password and rejected the wrong password.");
}
