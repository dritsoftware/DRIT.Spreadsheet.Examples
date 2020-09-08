using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DRIT.Spreadsheet;

namespace Filter
{
    class Program
    {
        static void Main(string[] args)
        {
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];
            worksheet.Name = "Values";
            worksheet.Columns[0].WidthCharacters = 15;

            worksheet["A1"].Value = "Column A";
            worksheet["A2"].Value = "ABC";
            worksheet["A3"].Value = "BCD";
            worksheet["A4"].Value = "CDE";
            worksheet["A5"].Value = "DEF";

            worksheet["A7"].Value = "ABC";
            worksheet["A8"].Value = "BCD";

            worksheet.Filter("A1:A8").Values(0, true, "ABC", "DEF");

            workbook.SaveAs(@"..\Out\Filter.xlsx");
        }
    }
}
