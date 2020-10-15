using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Office.Drawing;
using DRIT.Spreadsheet.Office.Model;

namespace Pictures
{
    class Program
    {
        static void Main(string[] args)
        {
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];

            worksheet.Columns["A"].WidthPixels = 120;
            worksheet.Columns["B"].WidthPixels = 90;
            worksheet.GetRange("A1:A5").SetRowsHeight(60);



            var picture1 = worksheet.Shapes.AddPicture(@"..\In\smiley.png", "B1",
                new Position(5, 5, ScreenMeasurementUnit.Pixel), new Size(0.75, 0.75, ScreenMeasurementUnit.Inch));

            worksheet["A2"].Value = "Brightness -50%";
            var picture2 = worksheet.Shapes.AddPicture(@"..\In\smiley.png", "B2",
                new Position(5, 5, ScreenMeasurementUnit.Pixel), new Size(0.75, 0.75, ScreenMeasurementUnit.Inch));
            picture2.Picture.BrightnessCorrection = -0.5m;

            worksheet["A3"].Value = "Contrast +80%";
            var picture3 = worksheet.Shapes.AddPicture(@"..\In\smiley.png", "B3",
                new Position(5, 5, ScreenMeasurementUnit.Pixel), new Size(0.75, 0.75, ScreenMeasurementUnit.Inch));
            picture3.Picture.ContrastCorrection = 0.8m;

            worksheet["A4"].Value = "Recolor: Sepia";
            var picture4 = worksheet.Shapes.AddPicture(@"..\In\smiley.png", "B4",
                new Position(5, 5, ScreenMeasurementUnit.Pixel), new Size(0.75, 0.75, ScreenMeasurementUnit.Inch));
            picture4.Picture.Recolor = Recolor.Sepia;

            workbook.SaveAs(@"..\Out\Pictures.xlsx");
        }
    }
}
