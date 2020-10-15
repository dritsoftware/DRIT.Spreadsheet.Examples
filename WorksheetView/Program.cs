using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Draw;

namespace WorksheetView
{
    class Program
    {
        static void Main(string[] args)
        {
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];
            worksheet.Name = "View1";

            worksheet.Columns["A"].WidthPixels = 250;

            worksheet["A1"].Value = "Show formulas";
            worksheet["A2"].Formula = "=SUM(1)";
            worksheet["A2"].Value = "Do not show row and column headers";
            worksheet["A3"].Value = "Zoom 125%";
            worksheet["A4"].Value = "Tab color: Red";
            worksheet["A5"].Value = "Gridline color: Green";

            worksheet.View.ShowFormulas = true;
            worksheet.View.ShowRowColumnHeaders = false;
            worksheet.View.ZoomScale = 125;
            worksheet.Properties.TabColor = SpreadsheetColor.Red;
            worksheet.View.GridlineColor = SpreadsheetColor.Green;

            var worksheet2 = workbook.AddWorksheet("View2");
            worksheet2.Columns["A"].WidthPixels = 250;
            worksheet2["A1"].Value = "Do not show zeros";
            worksheet2["B1"].Value = 0;
            worksheet2["A2"].Value = "Do not show gridlines";
            worksheet2["A3"].Value = "Page Layout View Zoom 75%";
            worksheet2["A4"].Value = "Page Layout View";

            worksheet2.View.ShowZeros = false;
            worksheet2.View.ShowGridLines = false;
            worksheet2.View.ZoomScalePageLayoutView = 75;
            worksheet2.View.ViewType = SheetViewType.PageLayout;

            var worksheet3 = workbook.AddWorksheet("Right to Left");
            worksheet3.View.RightToLeft = true;

            workbook.SaveAs(@"..\Out\WorksheetView.xlsx");
        }
    }
}
