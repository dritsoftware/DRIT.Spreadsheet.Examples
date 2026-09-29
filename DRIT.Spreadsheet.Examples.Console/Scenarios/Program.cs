using System.Linq;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Examples.Shared;

var workbook = new Workbook();
var worksheet = workbook.Worksheets[0];
worksheet["A1"].Value = "Revenue";
worksheet["B1"].Value = "Cost";
worksheet["A2"].Value = 100;
worksheet["B2"].Value = 60;

var baseline = new Scenario
{
    Name = "Baseline",
    User = "DRIT Examples",
    Locked = true,
    Comment = "Current forecast"
};
baseline.InputCells.Add(new InputCells { Cell = worksheet["A2"], Value = "100" });
baseline.InputCells.Add(new InputCells { Cell = worksheet["B2"], Value = "60" });
worksheet.Scenarios.Add(baseline);
worksheet.Scenarios.CurrentScenario = 0;
worksheet.Scenarios.LastShownScenario = 0;

var path = ExampleSupport.OutputPath("Scenarios.xlsx");
workbook.SaveAs(path);
var loaded = Workbook.Load(path);
var loadedScenario = loaded.Worksheets[0].Scenarios.Single();
ExampleSupport.Require(loadedScenario.Name == "Baseline", "The scenario name was not preserved.");
ExampleSupport.Require(loadedScenario.InputCells.Count == 2, "Scenario input cells were not preserved.");
ExampleSupport.Require(loadedScenario.InputCells.Any(cell => cell.Reference == "A2" && cell.Value == "100"), "The first scenario input was not preserved.");
ExampleSupport.Require(loaded.Worksheets[0].Scenarios.CurrentScenario == 0, "The current scenario was not preserved.");
Console.WriteLine("Scenario metadata and input cells preserved after reload.");
