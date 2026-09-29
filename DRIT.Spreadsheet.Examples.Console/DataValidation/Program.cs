using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
worksheet["A1"].Value = "Status";
worksheet["B1"].Value = "Score";
worksheet["D1"].Value = "Open";
worksheet["D2"].Value = "Closed";
worksheet.DataValidations.AddList("A2:A20", "=D1:D2");
worksheet.DataValidations.AddDecimal("B2:B20", DataValidationOperator.Between, 0, 100);
var path = ExampleSupport.OutputPath("DataValidation.xlsx");
workbook.SaveAs(path);

var loaded = new Workbook(path);
var validations = loaded.Worksheets[0].DataValidations;
ExampleSupport.Require(validations.Count == 2, "Both validations should be preserved.");
ExampleSupport.Require(validations[0].Type == DataValidationType.List, "The list validation type was not preserved.");
ExampleSupport.Require(validations[1].Operator == DataValidationOperator.Between, "The decimal validation operator was not preserved.");
ExampleSupport.ReportRoundTrip(path, loaded);
Console.WriteLine($"Validation rules: {validations.Count}");