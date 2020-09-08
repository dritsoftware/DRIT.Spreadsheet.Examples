using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DRIT.Spreadsheet;
using DRIT.Spreadsheet.Office.Chart;
using DRIT.Spreadsheet.Office.Drawing;
using DRIT.Spreadsheet.Office.Model;

namespace Charts
{
    class Program
    {
        static void Main(string[] args)
        {
            var workbook = new Workbook();

            var worksheet = workbook.Worksheets[0];
            worksheet.Name = "Data";
            //var sheet = workbook.AddWorksheet("SourceData");

            worksheet.GetRange("A1:D1").SetColumnsWidthPixel(110);

            worksheet.GetRange("B1:D1").SetValue(new[] { "Series 1", "Series 2", "Series 3" });
            worksheet.GetRange("A2:A5").SetValue(new[] { "Point 1", "Point 2", "Point 3", "Point 4" });
            worksheet.GetRange("B2:B5").SetValue(new[] { 2, 4, 6, 8 });
            worksheet.GetRange("C2:C5").SetValue(new[] { 4, 3, 2, 1 });
            worksheet.GetRange("D2:D5").SetValue(new[] { 2, 1, -1, -2 });

            TwoDArea(workbook);
            ThreeDArea(workbook);
            TwoDBar(workbook);
            TwoDLine(workbook);
            ThreeDLine(workbook);

            workbook.SaveAs(@"..\Out\ChartTypes.xlsx");
        }

        public static void TwoDArea(Workbook workbook)
        {
            var worksheet = workbook.AddWorksheet("2-D Area");

            var cellOffset = new PointDoubleUnit(19050, 19050, ScreenMeasurementUnit.Emu);
            worksheet.GetRange("A1:A4").SetRowsHeight(66);

            worksheet.GetRange("A8:A11").SetRowsHeight(66);

            var chart1 = worksheet.Charts.AddAreaChart(2, "Chart 1", "A1", cellOffset);
            chart1.Size.HeightInches = 2.5;
            chart1.Size.WidthInches = 3.5;
            chart1.DataSource = "Data!$A$1:$D$5";

            var chart2 = worksheet.Charts.AddAreaChart(3, "Chart 2", "A7", cellOffset);
            chart2.Size.HeightInches = 2.5;
            chart2.Size.WidthInches = 3.5;
            chart2.Grouping = Grouping.Stacked;
            chart2.DataSource = "Data!$A$1:$D$5";
        }

        public static void ThreeDArea(Workbook workbook)
        {
            var worksheet = workbook.AddWorksheet("3-D Area");

            var cellOffset = new PointDoubleUnit(19050, 19050, ScreenMeasurementUnit.Emu);
            worksheet.GetRange("A1:A4").SetRowsHeight(66);

            worksheet.GetRange("A8:A11").SetRowsHeight(66);

            var chart1 = worksheet.Charts.AddArea3DChart(2, "Chart 1", "A1", cellOffset);
            chart1.Size.HeightInches = 2.5;
            chart1.Size.WidthInches = 3.5;
            chart1.DataSource = "Data!$A$1:$D$5";

            var chart2 = worksheet.Charts.AddArea3DChart(3, "Chart 2", "A7", cellOffset);
            chart2.Size.HeightInches = 2.5;
            chart2.Size.WidthInches = 3.5;
            chart2.Grouping = Grouping.Stacked;
            chart2.DataSource = "Data!$A$1:$D$5";
        }

        public static void TwoDBar(Workbook workbook)
        {
            var worksheet = workbook.AddWorksheet("2-D Bar");

            var cellOffset = new PointDoubleUnit(19050, 19050, ScreenMeasurementUnit.Emu);
            worksheet.GetRange("A1:A4").SetRowsHeight(66);

            worksheet.GetRange("A8:A11").SetRowsHeight(66);

            var chart1 = worksheet.Charts.AddBarChart(2, "Chart 1", "A1", cellOffset);
            chart1.Grouping = BarGrouping.Clustered;
            chart1.Size.HeightInches = 2.5;
            chart1.Size.WidthInches = 3.5;
            chart1.DataSource = "Data!$A$1:$D$5";

            var chart2 = worksheet.Charts.AddBarChart(3, "Chart 2", "A7", cellOffset);
            chart2.Grouping = BarGrouping.Stacked;
            chart2.Size.HeightInches = 2.5;
            chart2.Size.WidthInches = 3.5;
            chart2.DataSource = "Data!$A$1:$D$5";
        }

        

        public static void TwoDLine(Workbook workbook)
        {
            var worksheet = workbook.AddWorksheet("2-D Line");

            var cellOffset = new PointDoubleUnit(19050, 19050, ScreenMeasurementUnit.Emu);
            worksheet.GetRange("A1:A4").SetRowsHeight(66);

            worksheet.GetRange("A8:A11").SetRowsHeight(66);

            var chart1 = worksheet.Charts.AddLineChart(2, "Chart 1", "A1", cellOffset);
            chart1.Size.HeightInches = 2.5;
            chart1.Size.WidthInches = 3.5;
            chart1.DataSource = "Data!$A$1:$D$5";

            var chart2 = worksheet.Charts.AddLineChart(3, "Chart 2", "A7", cellOffset);
            chart2.Size.HeightInches = 2.5;
            chart2.Size.WidthInches = 3.5;
            chart2.Grouping = Grouping.Stacked;
            chart2.DataSource = "Data!$A$1:$D$5";
        }

        public static void ThreeDLine(Workbook workbook)
        {
            var worksheet = workbook.AddWorksheet("3-D Line");

            var cellOffset = new PointDoubleUnit(19050, 19050, ScreenMeasurementUnit.Emu);
            worksheet.GetRange("A1:A4").SetRowsHeight(66);

            worksheet.GetRange("A8:A11").SetRowsHeight(66);

            var chart1 = worksheet.Charts.AddLine3DChart(2, "Chart 1", "A1", cellOffset);
            chart1.Size.HeightInches = 2.5;
            chart1.Size.WidthInches = 3.5;
            chart1.DataSource = "Data!$A$1:$D$5";

        }
    }
}
