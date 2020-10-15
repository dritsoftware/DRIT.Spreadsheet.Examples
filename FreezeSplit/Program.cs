using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DRIT.Spreadsheet;

namespace FreezeSplit
{
    class Program
    {
        static void Main(string[] args)
        {
            var workbook = new Workbook();
            var freezeSheet = workbook.Worksheets[0];
            freezeSheet.Name = "Freeze";

            //Freeze the first column and the top 2 rows
            freezeSheet.View.Panes.FreezePanes("B3");

            var splitSheet = workbook.AddWorksheet("Split");

            //Create four panes by splitting at F10
            splitSheet.View.Panes.SplitPanes("F10");
            workbook.SaveAs(@"..\Out\FreezeSplit.xlsx");
        }
    }
}
