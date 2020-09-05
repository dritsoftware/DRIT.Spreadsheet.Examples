using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Draw;
using DRIT.Spreadsheet.Office.Model;

namespace Formatting
{
    class Program
    {
        static void Main(string[] args)
        {
            var workbook = new Workbook();
            

            Fonts(workbook);
            Borders(workbook);
            Fills(workbook);
            NumberFormats(workbook);

            workbook.SaveAs(@"..\Out\Formatting.xlsx");
        }

        public static void Fonts(Workbook workbook)
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.Name = "Fonts";

            worksheet.Columns[0].WidthCharacters = 11;
            worksheet.Columns[1].WidthCharacters = 25;

            worksheet["A1"].Value = "Font Name:";
            worksheet["B1"].Value = "12 points";
            worksheet["B1"].Font.Size = 12;

            worksheet["A2"].Value = "Font Name:";
            worksheet["B2"].Value = "Arial";
            worksheet["B2"].Font.Name = "Arial";

            worksheet["B3"].Value = "Bold";
            worksheet["B3"].Font.Bold = true;

            worksheet["B4"].Value = "Italic";
            worksheet["B4"].Font.Italic = true;

            worksheet["B5"].Value = "Strikethrough";
            worksheet["B5"].Font.Strikethrough = true;

            worksheet["B6"].Value = "Subscript";
            worksheet["B6"].Font.Subscript = true;

            worksheet["B7"].Value = "Superscript";
            worksheet["B7"].Font.Superscript = true;

            worksheet["B8"].Value = "UnderLine Single";
            worksheet["B8"].Font.Underline = SpreadsheetUnderline.Single;

            worksheet["B9"].Value = "UnderLine Double Accounting";
            worksheet["B9"].Font.Underline = SpreadsheetUnderline.DoubleAccounting;

            worksheet["A10"].Value = "Font Color:";
            worksheet["B10"].Value = "Red";
            worksheet["B10"].Font.Color = SpreadsheetColor.Red;

            var paragraph = new RichTextParagraph();
            paragraph.Runs.Add(new RichTextRun("12 points ") { FontSize = 12 });
            paragraph.Runs.Add(new RichTextRun("Arial ") { FontName = "Arial" });
            paragraph.Runs.Add(new RichTextRun("bold ") { Bold = true });
            paragraph.Runs.Add(new RichTextRun("red") { Foreground = SpreadsheetColor.Red });
            worksheet["B11"].RichText = paragraph;
        }

        public static void Borders(Workbook workbook)
        {
            var worksheet = workbook.AddWorksheet("Borders");
            worksheet.Columns[0].WidthCharacters = 11;

            worksheet["A1"].Value = "Left Thick Border";
            worksheet["B1"].Borders.Left = new Border(BorderStyle.Thick);

            worksheet["A3"].Value = "Right Thick Border";
            worksheet["B3"].Borders.Right = new Border(BorderStyle.Thick);

            worksheet["A5"].Value = "Top Thick Border";
            worksheet["B5"].Borders.Top = new Border(BorderStyle.Thick);

            worksheet["A7"].Value = "Bottom Thick Border";
            worksheet["B7"].Borders.Bottom = new Border(BorderStyle.Thick);

            worksheet["A9"].Value = "Diagonal Up Border";
            worksheet["B9"].Borders.SetDiagonalUpBorder(BorderStyle.Thick);

            worksheet["A11"].Value = "Diagonal Down Border";
            worksheet["B11"].Borders.SetDiagonalDownBorder(BorderStyle.Thick);

            worksheet["A13"].Value = "Slant Dash Dot";
            worksheet["B13"].Borders.SetBoxBorder(BorderStyle.SlantDashDot);

            worksheet["A15"].Value = "Box Red";
            worksheet["B15"].Borders.SetBoxBorder(BorderStyle.Thick, SpreadsheetColor.Red);

            worksheet["A17"].Value = "Thin Cross Green";
            worksheet["B17"].Borders.SetDiagonalCrossBorder(BorderStyle.Thin, SpreadsheetColor.Green);
        }

        public static void Fills(Workbook workbook)
        {
            var worksheet = workbook.AddWorksheet("Fills");
            worksheet.Columns[0].WidthCharacters = 20;

            worksheet["A1"].Value = "Solid Blue";
            worksheet["B1"].Fill = CellFill.CreateSolidFill(SpreadsheetColor.Blue);

            worksheet["A3"].Value = "Green Red Dark Trellis";
            worksheet["B3"].Fill = CellFill.CreatePattern(SpreadsheetColor.Green, SpreadsheetColor.Red, PatternType.DarkTrellis);

            worksheet["A5"].Value = "Linear Gradient";
            worksheet["B5"].Fill = CellFill.CreateGradient(GradientType.Linear, 180, SpreadsheetColor.Blue, SpreadsheetColor.Red);
        }

        public static void NumberFormats(Workbook workbook)
        {
            var worksheet = workbook.AddWorksheet("NumberFormats");
            worksheet.Columns[0].WidthCharacters = 20;

            double doubleValue = 1234.5678;
            worksheet["A1"].Value = "General";
            worksheet["A2"].Value = "'0";
            worksheet["A3"].Value = "'0.00";
            worksheet["A4"].Value = "#,##0";
            worksheet["A5"].Value = "#,##0.00";

            worksheet["B1"].Value = doubleValue;
            worksheet["B2"].Value = doubleValue;
            worksheet["B3"].Value = doubleValue;
            worksheet["B4"].Value = doubleValue;
            worksheet["B5"].Value = doubleValue;

            worksheet["B1"].Format = DRIT.Spreadsheet.Office.Model.Formats.General;
            worksheet["B2"].Format = new NumberFormat("0");
            worksheet["B3"].Format = new NumberFormat("0.00");
            worksheet["B4"].Format = new NumberFormat("#,##0");
            worksheet["B5"].Format = new NumberFormat("#,##0.00");

            worksheet["A7"].Value = "#,##0.00;[Red]-#,##0.00";
            worksheet["B7"].Value = doubleValue * -1;
            worksheet["B7"].Format = new NumberFormat("#,##0.00;[Red]-#,##0.00");

            worksheet["A9"].Value = "Scientific 0.E+00";
            worksheet["B9"].Value = doubleValue;
            worksheet["B9"].Format = new NumberFormat("0.E+00");

            worksheet["A10"].Value = "Scientific 0.00E+0";
            worksheet["B10"].Value = doubleValue;
            worksheet["B10"].Format = new NumberFormat("0.00E+0");

            worksheet["A12"].Value = "Percent 0%";
            worksheet["B12"].Value = doubleValue;
            worksheet["B12"].Format = new NumberFormat("0%");

            worksheet["A13"].Value = "Scientific 0.00%";
            worksheet["B13"].Value = doubleValue;
            worksheet["B13"].Format = new NumberFormat("0.00%");
        }
    }
}
