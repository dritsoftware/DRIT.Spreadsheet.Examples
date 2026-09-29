using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

#if NET48
using System;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

using var rsa = RSA.Create(2048);
var request = new CertificateRequest(
    "CN=DRIT.Spreadsheet Phase 4 Example",
    rsa,
    HashAlgorithmName.SHA256,
    RSASignaturePadding.Pkcs1);
using var certificate = request.CreateSelfSigned(
    DateTimeOffset.UtcNow.AddDays(-1),
    DateTimeOffset.UtcNow.AddDays(30));

var workbook = new Workbook();
workbook.Worksheets[0]["A1"].Value = "Signed workbook";
var path = ExampleSupport.OutputPath("DigitalSignature.xlsx");
workbook.SaveAsSigned(path, new XlsxDigitalSignatureSaveOptions
{
    Certificate = certificate,
    SignerRole = "Phase 4 example"
});
var signatures = Workbook.ListDigitalSignatures(path);
ExampleSupport.Require(signatures.Count == 1, "The signed workbook should contain one digital signature.");
ExampleSupport.Require(!string.IsNullOrEmpty(signatures[0].SignaturePartUri), "The signature part was not recorded.");
Workbook.RemoveDigitalSignatures(path);
ExampleSupport.Require(Workbook.ListDigitalSignatures(path).Count == 0, "The digital signature was not removed.");
Console.WriteLine("Digital signature created, listed, and removed on net48.");
#else
Console.WriteLine("Digital signature example is platform-gated to net48 because OPC signing is unavailable on this target.");
#endif
