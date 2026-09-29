using System.IO;
using System.Linq;
using DRIT.Drawing.Imaging;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;
using DRIT.Spreadsheet.Export.Image;

var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
worksheet.Name = "ImageReport";
worksheet["A1"].Value = "Image export";
worksheet["A2"].Value = 2026;
worksheet["B2"].Value = 42;

var outputPattern = ExampleSupport.OutputPath("ImageExport-{sheet}-{page}.png");
workbook.SaveAsImage(outputPattern, new WorkbookImageSaveOptions
{
    Format = ImageExportFormat.Png,
    Dpi = 96,
    ForceGridLines = true
});

var imagePath = Directory.GetFiles(ExampleSupport.RootDirectory + "\\Out", "ImageExport-*.png")
    .Single();
using var image = new Bitmap(imagePath);
ExampleSupport.Require(image.Width > 0 && image.Height > 0, "The exported image has invalid dimensions.");
ExampleSupport.Require(new FileInfo(imagePath).Length > 0, "The exported image is empty.");
Console.WriteLine($"Image saved at {image.Width}x{image.Height}: {imagePath}");
