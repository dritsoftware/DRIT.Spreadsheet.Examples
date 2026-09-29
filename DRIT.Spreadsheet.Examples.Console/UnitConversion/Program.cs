using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

const double points = 72;
var pixels = UnitConverter.ToPixelFromPoint(points);
var roundTripPoints = UnitConverter.ToPointFromPixel(pixels);
var inches = UnitConverter.Convert(ScreenMeasurementUnit.Centimeter, ScreenMeasurementUnit.Inch, 2.54);
ExampleSupport.Require(Math.Abs(roundTripPoints - points) < 0.001, "Point/pixel conversion did not round-trip.");
ExampleSupport.Require(Math.Abs(inches - 1) < 0.001, "Centimeter/inch conversion was incorrect.");

var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
worksheet.Rows[0].HeightPoints = 24;
var path = ExampleSupport.OutputPath("UnitConversion.xlsx");
workbook.SaveAs(path);
ExampleSupport.Require(Math.Abs(UnitConverter.ToPointFromPixel(worksheet.Rows[0].HeightPixels) - 24) < 0.25, "Row height conversion was inconsistent.");
Console.WriteLine($"72 points = {pixels:0.##} pixels = {roundTripPoints:0.##} points");