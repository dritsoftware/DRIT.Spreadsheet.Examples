using System.IO;
using System.Text;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
var xml = "<batch xmlns=\"urn:drit:spreadsheet:phase4\"><value>42</value></batch>";
var schema = Encoding.UTF8.GetBytes("urn:drit:spreadsheet:phase4");
var index = workbook.CustomXmlParts.Add(Encoding.UTF8.GetBytes(xml), schema);
var part = workbook.CustomXmlParts[index];
var id = Guid.NewGuid().ToString();
part.ID = id;
ExampleSupport.Require(workbook.CustomXmlParts.SelectByID(id) == part, "The custom XML part could not be selected by ID.");
ExampleSupport.Require(part.XmlContent.Contains("<value>42</value>"), "The custom XML content was not assigned.");

var path = ExampleSupport.OutputPath("CustomXml.xlsx");
workbook.SaveAs(path);
using var archive = System.IO.Compression.ZipFile.OpenRead(path);
var itemEntry = archive.Entries.FirstOrDefault(entry => entry.FullName.StartsWith("customXml/item", StringComparison.Ordinal));
ExampleSupport.Require(itemEntry != null, "The custom XML item part was not saved.");
using var reader = new StreamReader(itemEntry.Open());
ExampleSupport.Require(reader.ReadToEnd().Contains("<value>42</value>"), "The custom XML content was not saved.");
Console.WriteLine($"Custom XML part saved: {part.ID}");
