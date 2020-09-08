using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Draw;

namespace Comments
{
    class Program
    {
        static void Main(string[] args)
        {
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];

            worksheet.GetRange("A1:A10").SetRowsHeight(60);
            worksheet.Columns["A"].WidthPixels = 100;
            worksheet.Columns["B"].WidthPixels = 100;

            worksheet["A2"].AddComment("DR-IT", "Comment");
            FormatComment(worksheet, "A2");
            worksheet["A2"].Comment.Line.Color = Color.Lime;

            worksheet["A3"].AddComment("DR-IT", "Comment");
            FormatComment(worksheet, "A3");
            worksheet["A3"].Comment.Line.Color = Color.Red;
            worksheet["A3"].Comment.Line.Transparency = 0.5m;

            worksheet["A4"].AddComment("DR-IT", "Comment");
            FormatComment(worksheet, "A4");
            worksheet["A4"].Comment.Line.WidthPoints = 5;

            worksheet["A5"].AddComment("DR-IT", "Comment");
            FormatComment(worksheet, "A5");
            worksheet["A5"].Comment.Fill.Color = Color.Yellow;

            worksheet["A6"].AddComment("DR-IT", "Comment");
            FormatComment(worksheet, "A6");
            worksheet["A6"].Comment.Fill.Color = Color.Red;
            worksheet["A6"].Comment.Fill.Type = CommentFillEffect.Gradient;
            worksheet["A6"].Comment.Fill.Gradient.DarkenPercentage = 0.6m;


            var richText = new RichTextParagraph();
            richText.Runs.Add(new RichTextRun() { FontName = "Arial", FontSize = 10, Foreground = SpreadsheetColor.Red, Text = "Arial 10 Red" });
            richText.Runs.Add(new RichTextRun() { FontName = "Cambria", FontSize = 6, Foreground = SpreadsheetColor.Green, Text = "Cambria 6 Green" });
            worksheet["A7"].AddComment("DR-IT", richText);
            FormatComment(worksheet, "A7");


            workbook.SaveAs(@"..\Out\Comments.xlsx");
        }

        private static void FormatComment(Worksheet worksheet, string label)
        {
            worksheet[label].Comment.Size.WidthPixels = 80;
            worksheet[label].Comment.Size.HeightPixels = 50;
            //Anchor the comment shape one row above and one column to the left
            worksheet[label].Comment.AnchoredCell = worksheet[worksheet[label].Row.Index - 1, worksheet[label].Column.Index + 1];
            //Always show the comment
            worksheet[label].Comment.Visible = true;
        }
    }
}
