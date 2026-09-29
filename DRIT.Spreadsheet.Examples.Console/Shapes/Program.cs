using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;
using DRIT.Spreadsheet.Drawing;

namespace Shapes
{
    class Program
    {
        static void Main(string[] args)
        {
            var workbook = new Workbook();
            Line(workbook);
            Fill(workbook);
            Shadow(workbook);
            workbook.SaveAs(ExampleSupport.OutputPath("Shapes.xlsx"));
        }

        public static void Line(Workbook workbook)
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.Name = "Line";

            worksheet.Columns["B"].WidthPixels = 90;
            worksheet.GetRange("A1:A10").SetRowsHeight(53);

            Position cellOffsetPixel = new Position(5, 5, ScreenMeasurementUnit.Pixel);
            Size size = new Size(738000, 428625, ScreenMeasurementUnit.Emu);

            worksheet["A1"].Value = "Line Red";
            var rectangle1 = worksheet.Shapes.AddShape(ShapeType.Rectangle, "B1", cellOffsetPixel, size);
            rectangle1.Line.DrawingColor = DrawingColor.Red;
            rectangle1.Line.WidthPoints = 2;

            worksheet["A2"].Value = "Line Red 50% Transparent";
            var rectangle2 = worksheet.Shapes.AddShape(ShapeType.Rectangle, "B2", cellOffsetPixel, size);
            rectangle2.Line.DrawingColor = DrawingColor.Red;
            rectangle2.Line.WidthPoints = 2;
            rectangle2.Line.Transparency = 0.5m;

            worksheet["A1"].Value = "Line Red";
            var rectangle3 = worksheet.Shapes.AddShape(ShapeType.Rectangle, "B3", cellOffsetPixel, size);
            rectangle3.Line.Gradient = new GradientFill();
            rectangle3.Line.Gradient.GradientStops.Add(new DRIT.Spreadsheet.Drawing.GradientStop(0, DrawingColor.Text1));
            rectangle3.Line.Gradient.GradientStops.Add(new DRIT.Spreadsheet.Drawing.GradientStop(1, DrawingColor.Red));
            rectangle3.Line.WidthPoints = 2;

            var rectangle4 = worksheet.Shapes.AddShape(ShapeType.Rectangle, "B4", cellOffsetPixel, size);
            rectangle4.Line.DrawingColor = DrawingColor.Text1;
            rectangle4.Line.WidthPoints = 5;
            rectangle4.Line.CompoundType = CompoundLine.Triple;
        }

        public static void Fill(Workbook workbook)
        {
            var worksheet = workbook.AddWorksheet("Fill");

            worksheet.Columns["A"].WidthPixels = 150;
            worksheet.Columns["B"].WidthPixels = 90;
            worksheet.GetRange("A1:A10").SetRowsHeight(53);

            Position cellOffsetPixel = new Position(5, 5, ScreenMeasurementUnit.Pixel);
            Size size = new Size(738000, 428625, ScreenMeasurementUnit.Emu);

            worksheet["A1"].Value = "Fill Red";
            var rectangle1 = worksheet.Shapes.AddShape(ShapeType.Rectangle, "B1", cellOffsetPixel, size);
            rectangle1.Fill.SolidDrawingColor = DrawingColor.Red;

            worksheet["A2"].Value = "Fill Red 50% Transparent";
            var rectangle2 = worksheet.Shapes.AddShape(ShapeType.Rectangle, "B2", cellOffsetPixel, size);
            rectangle2.Fill.SolidDrawingColor = DrawingColor.Red;
            rectangle2.Fill.Transparency = 0.5m;

            var rectangle3 = worksheet.Shapes.AddShape(ShapeType.Rectangle, "B3", cellOffsetPixel, size);
            rectangle3.Fill.Gradient = new GradientFill();
            rectangle3.Fill.Gradient.GradientStops.Add(new DRIT.Spreadsheet.Drawing.GradientStop(0, DrawingColor.Text1));
            rectangle3.Fill.Gradient.GradientStops.Add(new DRIT.Spreadsheet.Drawing.GradientStop(0.5m, DrawingColor.Yellow));
            rectangle3.Fill.Gradient.GradientStops.Add(new DRIT.Spreadsheet.Drawing.GradientStop(1, DrawingColor.Red));

            var rectangle4 = worksheet.Shapes.AddShape(ShapeType.Rectangle, "B4", cellOffsetPixel, size);
            rectangle4.Fill.Gradient = GradientPresets.Horizon();

            var rectangle5 = worksheet.Shapes.AddShape(ShapeType.Rectangle, "B5", cellOffsetPixel, size);
            rectangle5.Fill.Type = FillType.PatternFill;
            rectangle5.Fill.Pattern = new PatternFill();
            rectangle5.Fill.Pattern.Preset = DRIT.Spreadsheet.Drawing.PatternType.DiagonalBrick;
            rectangle5.Fill.Pattern.BackgroundColor = DrawingColor.Background;
            rectangle5.Fill.Pattern.ForegroundColor = DrawingColor.Red;
        }

        public static void Shadow(Workbook workbook)
        {
            var worksheet = workbook.AddWorksheet("Shadow");

            worksheet.Columns["B"].WidthPixels = 90;
            worksheet.GetRange("A1:A10").SetRowsHeight(53);

            Position cellOffsetPixel = new Position(5, 5, ScreenMeasurementUnit.Pixel);
            Size size = new Size(738000, 428625, ScreenMeasurementUnit.Emu);

            worksheet["A1"].Value = "Outer Shadow";
            var rectangle1 = worksheet.Shapes.AddShape(ShapeType.Rectangle, "B1", cellOffsetPixel, size);
            rectangle1.Effects.Shadow.Type = ShadowType.Outer;
            rectangle1.Effects.Shadow.Color = DrawingColor.Black;
            rectangle1.Effects.Shadow.Transparency = 0.4m;
            rectangle1.Effects.Shadow.Size = 1.2m;
            rectangle1.Effects.Shadow.Blur = 4;
            rectangle1.Effects.Shadow.Angle = 45;
            rectangle1.Effects.Shadow.Distance = 6;

            var rectangle2 = worksheet.Shapes.AddShape(ShapeType.Rectangle, "B2", cellOffsetPixel, size);
            rectangle2.Effects.Shadow.Type = ShadowType.Outer;
            rectangle2.Effects.Shadow.Color = DrawingColor.Red;
            rectangle2.Effects.Shadow.Transparency = 0.2m;
            rectangle2.Effects.Shadow.Size = 1;
            rectangle2.Effects.Shadow.Blur = 4;
            rectangle2.Effects.Shadow.Angle = 45;
            rectangle2.Effects.Shadow.Distance = 6;

            worksheet["A3"].Value = "Inner Shadow";
            var rectangle3 = worksheet.Shapes.AddShape(ShapeType.Rectangle, "B3", cellOffsetPixel, size);
            rectangle3.Fill.SolidDrawingColor = DrawingColor.Window;
            rectangle3.Effects.Shadow.Type = ShadowType.Inner;
            rectangle3.Effects.Shadow.Color = DrawingColor.Black;
            rectangle3.Effects.Shadow.Transparency = 0.4m;
            rectangle3.Effects.Shadow.Blur = 5;
            rectangle3.Effects.Shadow.Angle = 135;
            rectangle3.Effects.Shadow.Distance = 6;
        }
    }
}
