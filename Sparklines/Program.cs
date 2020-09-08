using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Chart;

namespace Sparklines
{
    class Program
    {
        static void Main(string[] args)
        {
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];

            worksheet.GetRange("A1:E1").SetValue(new[] { 1, 3, 2, 5, 4 });
            worksheet.GetRange("A2:E2").SetValue(new[] { 1, 3, 2, 5, 4 });

            worksheet.SparkLineGroups.Add("Sheet1!A1:E1", "F1", SparkLineType.Line);
            worksheet.SparkLineGroups.Add("Sheet1!A2:E2", "F2", SparkLineType.Column);

            workbook.SaveAs(@"..\Out\Sparklines.xlsx");
        }
    }
}
