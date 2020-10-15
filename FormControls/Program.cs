using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DRIT.Spreadsheet;

namespace FormControls
{
    class Program
    {
        static void Main(string[] args)
        {
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];

            worksheet.Columns["B"].WidthPixels = 250;
            worksheet.GetRange("A1:A10").SetRowsHeight(53);

            worksheet["A3"].Value = "ABC";
            worksheet["A4"].Value = "BCD";
            worksheet["A5"].Value = "CDE";

            var control = worksheet.FormControls.AddComboBox("Drop Down 1", "B1");
            control.Size.WidthInches = 2;
            control.Size.HeightInches = 0.22;
            control.Position.XPixels = 5;
            control.Position.YPixels = 5;
            control.InputRangeLabel = "Sheet1!$A$3:$A$5";
            control.CellLinkLabel = "A1";

            workbook.SaveAs(@"..\Out\FormControls.xlsx");
        }
    }
}
