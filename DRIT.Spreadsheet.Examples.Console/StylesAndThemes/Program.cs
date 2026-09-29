using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Drawing;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
workbook.Theme.SetCustomTheme("Phase 3 Theme", new[]
{
    DrawingColor.White, DrawingColor.Black, DrawingColor.White, DrawingColor.DarkSlateGray,
    DrawingColor.CornflowerBlue, DrawingColor.Orange, DrawingColor.SeaGreen,
    DrawingColor.MediumPurple, DrawingColor.Goldenrod, DrawingColor.Teal,
    DrawingColor.Blue, DrawingColor.Purple
});

var worksheet = workbook.Worksheets[0];
worksheet.Name = "Styles";
worksheet.GetRange("A1:B3").SetValue(new object[,] {
    { "Metric", "Value" },
    { "Revenue", 125000 },
    { "Margin", 0.32 }
});

var titleStyle = new CellStyleTemplate
{
    FontBold = true,
    FontSize = 14,
    FontColor = SpreadsheetColor.Text1,
    FillBackgroundColor = SpreadsheetColor.Accent1,
    FillForegroundColor = SpreadsheetColor.Accent1,
    BorderStyle = BorderStyle.Thin,
    BorderColor = SpreadsheetColor.Text1,
    HorizontalAlignment = HorizontalCellAlignment.Center
};
worksheet["A1"].ApplyStyle(titleStyle);
worksheet["B1"].ApplyStyle(titleStyle);
worksheet["B2"].Format = new NumberFormat("$#,##0");
worksheet["B3"].Format = new NumberFormat("0.0%");

var path = ExampleSupport.OutputPath("StylesAndThemes.xlsx");
workbook.SaveAs(path);
var loaded = new Workbook(path);
ExampleSupport.Require(loaded.Theme.Name == "Phase 3 Theme", "The custom theme was not saved.");
ExampleSupport.Require(loaded.Worksheets[0]["A1"].Font.Bold, "The reusable title style was not saved.");
ExampleSupport.ReportRoundTrip(path, loaded);
Console.WriteLine($"Theme: {loaded.Theme.Name}");