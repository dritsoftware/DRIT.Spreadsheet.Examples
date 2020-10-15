using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Vba;

namespace Vba
{
    class Program
    {
        static void Main(string[] args)
        {
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];

            var module = workbook.VbaProject.Modules["ThisWorkbook"];

            var codeBuilder = new StringBuilder();
            codeBuilder.AppendLine("Public Sub ShowIndexedColors()");
            codeBuilder.AppendLine();
            codeBuilder.AppendLine("Dim i As Integer");
            codeBuilder.AppendLine("Dim j As Integer");
            codeBuilder.AppendLine();
            codeBuilder.AppendLine("For i = 1 To 7");
            codeBuilder.AppendLine("\tFor j = 1 To 8");
            codeBuilder.AppendLine("\t\tCells(i, j).Value = 8 * (i - 1) + j");
            codeBuilder.AppendLine("\t\tCells(i, j).Interior.ColorIndex = 8 * (i - 1) + j");
            codeBuilder.AppendLine("\tNext j");
            codeBuilder.AppendLine("Next i");
            codeBuilder.AppendLine();
            codeBuilder.AppendLine("End Sub");
            module.Code = codeBuilder.ToString();

            var control = worksheet.FormControls.AddButton("Button 1", "A13");
            control.Size.WidthInches = 2.43;
            control.Size.HeightInches = 0.27;
            control.Position.XPixels = 5;
            control.Position.YPixels = 5;
            control.Macro = "[0]!ThisWorkbook.ShowIndexedColors";
            control.Text = "Show Indexed Colors";

            workbook.SaveAs(@"..\Out\Vba.xlsm");
        }
    }
}
