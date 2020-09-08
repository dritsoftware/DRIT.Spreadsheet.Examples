using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DRIT.Spreadsheet;

namespace EcmaFormulas
{
    class Program
    {
        static void Main(string[] args)
        {
            var workbook = new Workbook();
            Engineering(workbook);
            Financial(workbook);
            Logical(workbook);
            Math(workbook);
            Statistical(workbook);
            workbook.SaveAs(@"..\Out\EcmaFormulas.xlsx");
        }

        public static void Engineering(Workbook workbook)
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.Name = "Engineering";


            worksheet["A1"].Value = "BESSELI";
            worksheet["B1"].Formula = "=BESSELI(-5.6,0)";
            
            worksheet["A2"].Value = "BESSELJ";
            worksheet["B2"].Formula = "=BESSELJ(-5.6,0)";
            worksheet["C2"].Formula = "=BESSELJ(2.345,5)";

            worksheet["A3"].Value = "BESSELK";
            worksheet["B3"].Formula = "=BESSELK(2.345,5)";

            worksheet["A4"].Value = "BESSELY";
            worksheet["B4"].Formula = "=BESSELY(2.345,5)";

            worksheet["A5"].Value = "BIN2DEC";
            worksheet["B5"].Formula = "=BIN2DEC(111)";
            worksheet["C5"].Formula = "=BIN2DEC(11111111)";
            worksheet["D5"].Formula = "=BIN2DEC(1111111110)";
            worksheet["E5"].Formula = "=BIN2DEC(1000000000)";

            worksheet["A6"].Value = "BIN2HEX";
            worksheet["B6"].Formula = "=BIN2HEX(1)";
            worksheet["C6"].Formula = "=BIN2HEX(1,4)";
            worksheet["D6"].Formula = "=BIN2HEX(111111)";
            worksheet["E6"].Formula = "=BIN2HEX(1111000000)";
            worksheet["F6"].Formula = "=BIN2HEX(1000000000,3)";

            worksheet["A7"].Value = "BIN2OCT";
            worksheet["B7"].Formula = "=BIN2OCT(1)";
            worksheet["C7"].Formula = "=BIN2OCT(1,4)";
            worksheet["D7"].Formula = "=BIN2OCT(111111)";
            worksheet["E7"].Formula = "=BIN2OCT(1111000000)";
            worksheet["F7"].Formula = "=BIN2OCT(1000000000,3)";
        }

        public static void Financial(Workbook workbook)
        {
            var worksheet = workbook.AddWorksheet("Financial");
            worksheet["A1"].Value = "ACCRINT";
            worksheet["B1"].Formula = "=ACCRINT(DATE(2006,3,1),DATE(2006,9,1),DATE(2006,5,1),0.1,1100,2,0)";
            worksheet["C1"].Formula = "=ACCRINT(DATE(2006,3,1),DATE(2006,9,1),DATE(2006,5,1),0.1,,2,0)";

            worksheet["A2"].Value = "ACCRINTM";
            worksheet["B2"].Formula = "=ACCRINTM(DATE(2006,3,1),DATE(2006,5,1),0.1,1100,0)";
            worksheet["C2"].Formula = "=ACCRINTM(DATE(2006,3,1),DATE(2006,5,1),0.1,,0)";
            worksheet["D2"].Formula = "=ACCRINTM(DATE(2006,3,1),DATE(2006,5,1),0.1,)";

            worksheet["A3"].Value = "AMORDEGRC";
            worksheet["B3"].Formula = "=AMORDEGRC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,1)";

            worksheet["A4"].Value = "AMORLINC";
            worksheet["B4"].Formula = "=AMORLINC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,1)";
        }

        public static void Logical(Workbook workbook)
        {
            var worksheet = workbook.AddWorksheet("Logical");

            worksheet["A1"].Value = "AND";
            worksheet["B1"].Formula = "=AND(TRUE)";
            worksheet["C1"].Formula = "=AND(TRUE,FALSE)";
            worksheet["D1"].Formula = "=AND(10>5,3=1+2,5)";
        }

        public static void Math(Workbook workbook)
        {
            var worksheet = workbook.AddWorksheet("Math");
            worksheet["A1"].Value = "ABS";
            worksheet["B1"].Formula = "=ABS(10.5)";
            worksheet["C1"].Formula = "=ABS(0)";
            worksheet["D1"].Formula = "=ABS(-10.5)";

            worksheet["A2"].Value = "ACOS";
            worksheet["B2"].Formula = "=ACOS(-1)";
            worksheet["C2"].Formula = "=ACOS(0)";
            worksheet["D2"].Formula = "=ACOS(1)";

            worksheet["A3"].Value = "ACOSH";
            worksheet["B3"].Formula = "=ACOSH(1)";
            worksheet["C3"].Formula = "=ACOSH(10)";
            worksheet["D3"].Formula = "=ACOSH(100)";

            worksheet["A4"].Value = "ASIN";
            worksheet["B4"].Formula = "=ASIN(-1)";
            worksheet["C4"].Formula = "=ASIN(0)";
            worksheet["D4"].Formula = "=ASIN(1)";

            worksheet["A5"].Value = "ASINH";
            worksheet["B5"].Formula = "=ASINH(10)";
            worksheet["C5"].Formula = "=ASINH(100)";
            worksheet["D5"].Formula = "=ASINH(0.5)";

            worksheet["A6"].Value = "ATAN";
            worksheet["B6"].Formula = "=ATAN(-1)";
            worksheet["C6"].Formula = "=ATAN(0)";
            worksheet["D6"].Formula = "=ATAN(1)";
            worksheet["E6"].Formula = "=ATAN(-10)";
            worksheet["F6"].Formula = "=ATAN(10)";

            worksheet["A7"].Value = "ATAN2";
            worksheet["B7"].Formula = "=ATAN2(1,1)";
            worksheet["C7"].Formula = "=ATAN2(-2,2)";
            worksheet["D7"].Formula = "=ATAN2(3,-3)";

            worksheet["A8"].Value = "ATANH";
            worksheet["B8"].Formula = "=ATANH(-0.999999)";
            worksheet["C8"].Formula = "=ATANH(0)";
            worksheet["D8"].Formula = "=ATANH(0.999999)";
        }

        public static void Statistical(Workbook workbook)
        {
            var worksheet = workbook.AddWorksheet("Statistical");

            worksheet["A1"].Value = "AVEDEV";
            worksheet["B1"].Formula = "=AVEDEV(-3.5,1.4,6.9,-4.5)";
            worksheet["C1"].Formula = "=AVEDEV({-3.5,1.4,6.9,-4.5})";

            worksheet["A2"].Value = "AVERAGE";
            worksheet["B2"].Formula = "=AVERAGE(1,2,3,4,5)";
            worksheet["C2"].Formula = "=AVERAGE({1,2;3,4})";
            worksheet["D2"].Formula = "=AVERAGE({1,2,3,4,5},6,\"7\")";
            worksheet["E2"].Formula = "=AVERAGE({1,\"2\",TRUE,4})";

            worksheet["F3"].Value = true;
            worksheet["G3"].Value = false;
            worksheet["A3"].Value = "AVERAGEA";
            worksheet["B3"].Formula = "=AVERAGEA(10,E3)";
            
            worksheet["C4"].Value = 10;
            worksheet["D4"].Value = 20;
            worksheet["E4"].Value = 30;
            worksheet["A4"].Value = "AVERAGEIF";
            worksheet["B4"].Formula = "=AVERAGEIF(C4:E4,\"15\")";

            worksheet["A5"].Value = "BETADIST";
            worksheet["B5"].Formula = "=BETADIST(0.5,1,2)";
            worksheet["C5"].Formula = "=BETADIST(0.5,1,2,-4.5,7.3)";
            worksheet["D5"].Formula = "=BETADIST(0.5,1,2,,2.3)";

            worksheet["A6"].Value = "BETAINV";
            worksheet["B6"].Formula = "=BETAINV(0.5,1,2)";
            worksheet["C6"].Formula = "=BETAINV(0.5,1,2,-4.5,7.3)";
            worksheet["D6"].Formula = "=BETAINV(0.5,1,2,,2.3)";

            worksheet["A7"].Value = "BINOMDIST";
            worksheet["B7"].Formula = "=BINOMDIST(6,10,0.5,FALSE)";
            worksheet["C7"].Formula = "=BINOMDIST(6,10,0.5,TRUE)";
        }
    }
}
