using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Drawing;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
worksheet.Name = "Anchors";
worksheet.Columns["B"].WidthPixels = 140;
worksheet.GetRange("A1:A8").SetRowsHeight(36);
var position = new Position(8, 6, ScreenMeasurementUnit.Pixel);
var size = new Size(120, 60, ScreenMeasurementUnit.Pixel);
var shape = worksheet.Shapes.AddShape(ShapeType.Rectangle, "B3", position, size);
shape.Name = "AnchoredCallout";
shape.Fill.SolidDrawingColor = DrawingColor.LightSkyBlue;
shape.Line.DrawingColor = DrawingColor.DarkBlue;
shape.Line.WidthPoints = 1.5;

var path = ExampleSupport.OutputPath("ShapeAnchoring.xlsx");
workbook.SaveAs(path);
ExampleSupport.Require(worksheet.Shapes.Count == 1, "The anchored shape was not added.");
ExampleSupport.Require(shape.ShapeType == ShapeType.Rectangle, "The anchored shape type was not preserved.");
ExampleSupport.Require(shape.Position != null && shape.Size != null, "The shape placement was not configured.");
Console.WriteLine("Shape anchored to B3 with pixel offset and size.");