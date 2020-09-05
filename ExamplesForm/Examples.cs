using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.ConditionalFormatting;
using DRIT.Spreadsheet.Draw;

namespace ExamplesForm
{
    public partial class Examples : Form
    {
        public Examples()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var workbook = new Workbook();
            CellValue(workbook);
            workbook.SaveAs(@"..\Out\ConditionalFormatting.xlsx");
        }

        public static void CellValue(Workbook workbook)
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.Name = "CellValue";
            worksheet.Columns[0].WidthCharacters = 40;
            worksheet.Columns.SetWidthCharacters("B", "F", 5);

            worksheet.GetRange("B1:F1").SetValue(new[] { 1, 2, 3, 4, 5 });
            worksheet.GetRange("B2:F2").SetValue(new[] { 1, 2, 3, 4, 5 });
            worksheet.GetRange("B3:F3").SetValue(new[] { 1, 2, 3, 4, 5 });
            worksheet.GetRange("B4:F4").SetValue(new[] { 1, 2, 3, 4, 5 });
            worksheet.GetRange("B5:F5").SetValue(new[] { 1, 2, 3, 4, 5 });
            worksheet.GetRange("B6:F6").SetValue(new[] { 1, 2, 3, 4, 5 });
            worksheet.GetRange("B7:F7").SetValue(new[] { 1, 2, 3, 4, 5 });
            worksheet.GetRange("B8:F8").SetValue(new[] { 1, 2, 3, 4, 5 });


            worksheet["A1"].Value = "Cell Value between 2 and 4";
            CreateCellValueConditionalFormatting(worksheet, "B1:F1", SpreadsheetColor.Accent1, ConditionalFormattingOperator.Between, 2, 4);

            worksheet["A2"].Value = "Cell Value not between 2 and 4";
            CreateCellValueConditionalFormatting(worksheet, "B2:F2", SpreadsheetColor.Accent1, ConditionalFormattingOperator.NotBetween, 2, 4);

            worksheet["A3"].Value = "Cell Value equal to 3";
            CreateCellValueConditionalFormatting(worksheet, "B3:F3", SpreadsheetColor.Accent1, ConditionalFormattingOperator.Equal, 3);

            worksheet["A4"].Value = "Cell Value not equal to 3";
            CreateCellValueConditionalFormatting(worksheet, "B4:F4", SpreadsheetColor.Accent1, ConditionalFormattingOperator.NotEqual, 3);

            worksheet["A5"].Value = "Cell Value greater than 3";
            CreateCellValueConditionalFormatting(worksheet, "B5:F5", SpreadsheetColor.Accent1, ConditionalFormattingOperator.GreaterThan, 3);

            worksheet["A6"].Value = "Cell Value less than 3";
            CreateCellValueConditionalFormatting(worksheet, "B6:F6", SpreadsheetColor.Accent1, ConditionalFormattingOperator.LessThan, 3);

            worksheet["A7"].Value = "Cell Value greater than or equal to 3";
            CreateCellValueConditionalFormatting(worksheet, "B7:F7", SpreadsheetColor.Accent1, ConditionalFormattingOperator.GreaterThanOrEqual, 3);

            worksheet["A8"].Value = "Cell Value less than or equal to 3";
            CreateCellValueConditionalFormatting(worksheet, "B8:F8", SpreadsheetColor.Accent1, ConditionalFormattingOperator.LessThanOrEqual, 3);
        }

        internal static void CreateCellValueConditionalFormatting(Worksheet worksheet, string range, SpreadsheetColor color, ConditionalFormattingOperator conditionalFormattingOperator, params double[] values)
        {
            var formatting = worksheet.ConditionalFormattings.Add(range);
            var rule = new CellIsFormattingRule();
            rule.Fill = CellFill.CreateSolidFill(color);
            rule.Operator = conditionalFormattingOperator;
            rule.Formula1 = values[0].ToString(CultureInfo.InvariantCulture);
            if (values.Length == 2)
            {
                rule.Formula2 = values[1].ToString(CultureInfo.InvariantCulture);
            }

            formatting.Rules.Add(rule);
        }
    }
}
