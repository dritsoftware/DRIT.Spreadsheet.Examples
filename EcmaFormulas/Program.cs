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
            DateTime(workbook);
            Engineering(workbook);
            Financial(workbook);
            Information(workbook);
            Logical(workbook);
            LookupReference(workbook);
            Math(workbook);
            Statistical(workbook);
            Text(workbook);
            workbook.SaveAs(@"..\Out\EcmaFormulas.xlsx");
        }

        public static void DateTime(Workbook workbook)
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.Name = "DateTime";

            worksheet["A1"].Value = "DATE";
            worksheet["B1"].Formula = "=DATE(0,1,1)";
            worksheet["C1"].Formula = "=DATE(1899,1,1)";
            worksheet["D1"].Formula = "=DATE(1900,1,1)";
            worksheet["E1"].Formula = "=DATE(9999,12,31)";

            worksheet["A2"].Value = "DATEDIF";
            worksheet["B2"].Formula = "=DATEDIF(DATE(2001,1,1),DATE(2003,1,1),\"Y\")";
            worksheet["C2"].Formula = "=DATEDIF(DATE(2001,6,1),DATE(2002,8,15),\"D\")";
            worksheet["D2"].Formula = "=DATEDIF(DATE(2001,6,1),DATE(2002,8,15),\"YD\")";
            worksheet["E2"].Formula = "=DATEDIF(DATE(2001,6,1),DATE(2002,8,15),\"MD\")";

            worksheet["A3"].Value = "DATEVALUE";
            worksheet["B3"].Formula = "=DATEVALUE(\"2/1/2006\")";
            worksheet["C3"].Formula = "=DATEVALUE(\"01-Feb-2006 10:06 AM\")";
            worksheet["D3"].Formula = "=DATEVALUE(\"2006/2/1\")";
            worksheet["E3"].Formula = "=DATEVALUE(\"2006-2-1\")";
            worksheet["F3"].Formula = "=DATEVALUE(\"1-Feb\")";

            worksheet["A4"].Value = "DAY";
            worksheet["B4"].Formula = "=DAY(DATE(2006,1,2))";
            worksheet["C4"].Formula = "=DAY(DATE(2006,0,2))";
            worksheet["D4"].Formula = "=DAY(DATE(2013,9,0))";
            worksheet["E4"].Formula = "=DAY(\"2006/1/2 10:45 AM\")";
            worksheet["F4"].Formula = "=DAY(30000)";

            worksheet["A5"].Value = "DAYS360";
            worksheet["B5"].Formula = "=DAYS360(DATE(2002,2,3),DATE(2005,5,31))";
            worksheet["C5"].Formula = "=DAYS360(DATE(2005,5,31),DATE(2002,2,3))";
            worksheet["D5"].Formula = "=DAYS360(DATE(2002,2,3),DATE(2005,5,31),FALSE)";
            worksheet["E5"].Formula = "=DAYS360(DATE(2002,2,3),DATE(2005,5,31),TRUE)";

            worksheet["A6"].Value = "EDATE";
            worksheet["B6"].Formula = "=EDATE(DATE(2006,1,31),5)";
            worksheet["C6"].Formula = "=EDATE(DATE(2004,2,29),12)";
            worksheet["D6"].Formula = "=EDATE(DATE(2004,2,28),12)";
            worksheet["E6"].Formula = "=EDATE(DATE(2004,1,15),-23)";

            worksheet["A7"].Value = "EOMONTH";
            worksheet["B7"].Formula = "=EOMONTH(DATE(2006,1,31),5)";
            worksheet["C7"].Formula = "=EOMONTH(DATE(2004,2,29),12)";
            worksheet["D7"].Formula = "=EOMONTH(DATE(2004,2,28),12)";
            worksheet["E7"].Formula = "=EOMONTH(DATE(2004,1,15),-23)";

            worksheet["A8"].Value = "HOUR";
            worksheet["B8"].Formula = "=HOUR(DATE(2006,2,26)+TIME(2,10,20))";
            worksheet["C8"].Formula = "=HOUR(TIME(22,56,34))";
            worksheet["D8"].Formula = "=HOUR(0)";
            worksheet["E8"].Formula = "=HOUR(10.5)";
            worksheet["F8"].Formula = "=HOUR(\"22-Oct-2001 10:53:12\")";
            worksheet["G8"].Formula = "=HOUR(\"10:53:12 pm\")";
            worksheet["H8"].Formula = "=HOUR(\"22:53:12\")";
        }

        public static void Engineering(Workbook workbook)
        {
            var worksheet = workbook.AddWorksheet("Engineering");

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

            worksheet["A8"].Value = "COMPLEX";
            worksheet["B8"].Formula = "=COMPLEX(-3.5,19.6)";
            worksheet["C8"].Formula = "=COMPLEX(3.5,-19.6,\"j\")";
            worksheet["D8"].Formula = "=COMPLEX(3.5,0)";
            worksheet["E8"].Formula = "=COMPLEX(0,2.4)";
            worksheet["F8"].Formula = "=COMPLEX(0,0)";

            worksheet["A8"].Value = "CONVERT";
            worksheet["B8"].Formula = "=CONVERT(10,\"ozm\",\"g\")";
            worksheet["C8"].Formula = "=CONVERT(1,\"yd\",\"mm\")";
            worksheet["D8"].Formula = "=CONVERT(1,\"yd\",\"cm\")";
            worksheet["E8"].Formula = "=CONVERT(1,\"yd\",\"m\")";
            worksheet["F8"].Formula = "=CONVERT(1,\"yd\",\"km\")";
            worksheet["G8"].Formula = "=CONVERT(1,\"mi\",\"Nmi\")";
            worksheet["H8"].Formula = "=CONVERT(1,\"day\",\"sec\")";
            worksheet["I8"].Formula = "=CONVERT(0,\"K\",\"C\")";

            worksheet["A9"].Value = "DEC2BIN";
            worksheet["B9"].Formula = "=DEC2BIN(23)";
            worksheet["C9"].Formula = "=DEC2BIN(-256)";
            worksheet["D9"].Formula = "=DEC2BIN(18,7)";

            worksheet["A10"].Value = "DEC2HEX";
            worksheet["B10"].Formula = "=DEC2HEX(23)";
            worksheet["C10"].Formula = "=DEC2HEX(-256)";
            worksheet["D10"].Formula = "=DEC2HEX(18,7)";

            worksheet["A11"].Value = "DEC2OCT";
            worksheet["B11"].Formula = "=DEC2OCT(23)";
            worksheet["C11"].Formula = "=DEC2OCT(-256)";
            worksheet["D11"].Formula = "=DEC2OCT(18,7)";

            worksheet["A12"].Value = "DELTA";
            worksheet["B12"].Formula = "=DELTA(10.5,10.5)";
            worksheet["C12"].Formula = "=DELTA(10.5,10.6)";
            worksheet["D12"].Formula = "=DELTA(10.5)";
            worksheet["E12"].Formula = "=DELTA(0)";

            worksheet["A13"].Value = "ERF";
            worksheet["B13"].Formula = "=ERF(1.234,4.5432)";
            worksheet["C13"].Formula = "=ERF(0,1.345)";
            worksheet["D13"].Formula = "=ERF(0,1.345)";

            worksheet["A14"].Value = "ERFC";
            worksheet["B14"].Formula = "=ERFC(1.234)";
            worksheet["C14"].Formula = "=ERFC(0)";

            worksheet["A15"].Value = "GESTEP";
            worksheet["B15"].Formula = "=GESTEP(5.6,-4.3)";
            worksheet["C15"].Formula = "=GESTEP(5.6,5.6)";
            worksheet["D15"].Formula = "=GESTEP(-5.6)";

            worksheet["A16"].Value = "HEX2BIN";
            worksheet["B16"].Formula = "=HEX2BIN(\"fE\")";
            worksheet["C16"].Formula = "=HEX2BIN(\"FFFFFFFFFE\")";
            worksheet["D16"].Formula = "=HEX2BIN(\"2\")";
            worksheet["D16"].Formula = "=HEX2BIN(\"F\",6)";

            worksheet["A17"].Value = "HEX2DEC";
            worksheet["B17"].Formula = "=HEX2DEC(\"fE\")";
            worksheet["C17"].Formula = "=HEX2DEC(\"FFFFFFFFFE\")";
            worksheet["D17"].Formula = "=HEX2DEC(\"F000000000\")";

            worksheet["A18"].Value = "HEX2OCT";
            worksheet["B18"].Formula = "=HEX2OCT(\"fE\")";
            worksheet["C18"].Formula = "=HEX2OCT(\"FFFFFFFFFE\")";
            worksheet["D18"].Formula = "=HEX2OCT(\"2\")";
            worksheet["D18"].Formula = "=HEX2OCT(\"F\",6)";

            worksheet["A19"].Value = "IMABS";
            worksheet["B19"].Formula = "=IMABS(\"3+4i\")";
            worksheet["C19"].Formula = "=IMABS(\"-2.5-34.6j\")";

            worksheet["A20"].Value = "IMAGINARY";
            worksheet["B20"].Formula = "=IMAGINARY(\"3+4i\")";
            worksheet["C20"].Formula = "=IMAGINARY(\"-2.5-34.6j\")";

            worksheet["A21"].Value = "IMARGUMENT";
            worksheet["B21"].Formula = "=IMARGUMENT(\"13+4i\")";
            worksheet["C21"].Formula = "=IMARGUMENT(\"-2.5-5j\")";

            worksheet["A22"].Value = "IMCONJUGATE";
            worksheet["B22"].Formula = "=IMCONJUGATE(\"2.3+4.5i\")";
            worksheet["C22"].Formula = "=IMCONJUGATE(\"-1-4j\")";

            worksheet["A23"].Value = "IMCOS";
            worksheet["B23"].Formula = "=IMCOS(\"2.3+4.5i\")";
            worksheet["C23"].Formula = "=IMCOS(\"-1-4j\")";

            worksheet["A24"].Value = "IMDIV";
            worksheet["B24"].Formula = "=IMDIV(\"13+4i\",\"5+3i\")";
            worksheet["C24"].Formula = "=IMDIV(\"-3-3.5i\",\"5+3i\")";

            worksheet["A25"].Value = "IMEXP";
            worksheet["B25"].Formula = "=IMEXP(\"2.3+4.5i\")";
            worksheet["C25"].Formula = "=IMEXP(\"-1-4j\")";

            worksheet["A26"].Value = "IMLN";
            worksheet["B26"].Formula = "=IMLN(\"3+4i\")";
            worksheet["C26"].Formula = "=IMLN(\"-2.5-34.6j\")";

            worksheet["A27"].Value = "IMLOG10";
            worksheet["B27"].Formula = "=IMLOG10(\"3+4i\")";
            worksheet["C27"].Formula = "=IMLOG10(\"-2.5-34.6j\")";

            worksheet["A28"].Value = "IMLOG2";
            worksheet["B28"].Formula = "=IMLOG2(\"3+4i\")";
            worksheet["C28"].Formula = "=IMLOG2(\"-2.5-34.6j\")";

            worksheet["A29"].Value = "IMPOWER";
            worksheet["B29"].Formula = "=IMPOWER(\"2.3+4.5i\",2.5)";
            worksheet["C29"].Formula = "=IMPOWER(\"-1-4j\",-3.56)";

            worksheet["A30"].Value = "IMPRODUCT";
            worksheet["B30"].Formula = "=IMPRODUCT(\"13+4i\")";
            worksheet["C30"].Formula = "=IMPRODUCT(\"-3-3.5i\",\"5+3i\")";
            worksheet["D30"].Formula = "=IMPRODUCT(\"1.3-2j\",\"-3.4+3j\",\"2.3-6j\")";

            worksheet["A31"].Value = "IMREAL";
            worksheet["B31"].Formula = "=IMREAL(\"3+4i\")";
            worksheet["C31"].Formula = "=IMREAL(\"-2.5-34.6j\")";

            worksheet["A32"].Value = "IMSIN";
            worksheet["B32"].Formula = "=IMSIN(\"2.3+4.5i\")";
            worksheet["C32"].Formula = "=IMSIN(\"-1-4j\")";

            worksheet["A33"].Value = "IMSQRT";
            worksheet["B33"].Formula = "=IMSQRT(\"2.3+4.5i\")";
            worksheet["C33"].Formula = "=IMSQRT(\"-1-4j\")";

            worksheet["A34"].Value = "IMSUM";
            worksheet["B34"].Formula = "=IMSUM(\"3+4i\")";
            worksheet["C34"].Formula = "=IMSUM(\"3+4i\",\"5-3i\")";
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

            worksheet["A5"].Value = "COUPDAYBS";
            worksheet["B5"].Formula = "=COUPDAYBS(DATE(2007,1,25),DATE(2008,11,15),2,1)";
            worksheet["C5"].Formula = "=COUPDAYBS(DATE(2007,1,25),DATE(2008,11,15),2)";

            worksheet["A6"].Value = "COUPDAYS";
            worksheet["B6"].Formula = "=COUPDAYS(DATE(2007,1,25),DATE(2008,11,15),2,1)";
            worksheet["C6"].Formula = "=COUPDAYS(DATE(2007,1,25),DATE(2008,11,15),2)";

            worksheet["A7"].Value = "COUPDAYSNC";
            worksheet["B7"].Formula = "=COUPDAYSNC(DATE(2007,1,25),DATE(2008,11,15),2,1)";
            worksheet["C7"].Formula = "=COUPDAYSNC(DATE(2007,1,25),DATE(2008,11,15),2)";

            worksheet["A8"].Value = "COUPNCD";
            worksheet["B8"].Formula = "=COUPNCD(DATE(2007,1,25),DATE(2008,11,15),2,1)";

            worksheet["A9"].Value = "COUPNUM";
            worksheet["B9"].Formula = "=COUPNUM(DATE(2007,1,25),DATE(2008,11,15),2,1)";

            worksheet["A10"].Value = "COUPPCD";
            worksheet["B10"].Formula = "=COUPPCD(DATE(2007,1,25),DATE(2008,11,15),2,1)";

            worksheet["A11"].Value = "CUMIPMT";
            worksheet["B11"].Formula = "=CUMIPMT(0.09/12,30*12,125000,13,24,0)";
            worksheet["C11"].Formula = "=CUMIPMT(0.09/12,30*12,125000,1,1,0)";

            worksheet["A12"].Value = "CUMPRINC";
            worksheet["B12"].Formula = "=CUMPRINC(0.09/12,30*12,125000,13,24,0)";
            worksheet["C12"].Formula = "=CUMPRINC(0.09/12,30*12,125000,1,1,0)";

            worksheet["A13"].Value = "DB";
            worksheet["B13"].Formula = "=DB(1000000,100000,6,1,7)";
            worksheet["C13"].Formula = "=DB(1000000,100000,6,2,7)";
            worksheet["D13"].Formula = "=DB(1000000,100000,6,7,7)";

            worksheet["A14"].Value = "DDB";
            worksheet["B14"].Formula = "=DDB(2400,300,10*365,1)";
            worksheet["C14"].Formula = "=DDB(2400,300,10*12,1,2)";
            worksheet["D14"].Formula = "=DDB(2400,300,10,1,2)";
            worksheet["E14"].Formula = "=DDB(2400,300,10,2,1.5)";
            worksheet["F14"].Formula = "=DDB(2400,300,10,10)";

            worksheet["A15"].Value = "DISC";
            worksheet["B15"].Formula = "=DISC(DATE(2007,1,25),DATE(2007,6,15),97.975,100,1)";

            worksheet["A16"].Value = "DOLLARDE";
            worksheet["B16"].Formula = "=DOLLARDE(1.02,16)";
            worksheet["C16"].Formula = "=DOLLARDE(1.1,32)";

            worksheet["A17"].Value = "DOLLARFR";
            worksheet["B17"].Formula = "=DOLLARFR(1.125,16)";
            worksheet["C17"].Formula = "=DOLLARFR(1.125,32)";

            worksheet["A18"].Value = "DURATION";
            worksheet["B18"].Formula = "=DURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)";

            worksheet["A19"].Value = "EFFECT";
            worksheet["B19"].Formula = "=EFFECT(0.0525,4)";

            worksheet["A20"].Value = "FV";
            worksheet["B20"].Formula = "=FV(0.06/12,10,-200,-500,1)";
            worksheet["C20"].Formula = "=FV(0.12/12,12,-1000)";
            worksheet["D20"].Formula = "=FV(0.11/12,35,-2000,,1)";
            worksheet["E20"].Formula = "=FV(0.06/12,12,-100,-1000,1)";

            worksheet["A21"].Value = "FVSCHEDULE";
            worksheet["B21"].Formula = "=FVSCHEDULE(1,{0.09,0.11,0.1})";

            worksheet["A22"].Value = "INTRATE";
            worksheet["B22"].Formula = "=INTRATE(DATE(2008,2,15),DATE(2008,5,15),1000000,1014420,2)";

            worksheet["A23"].Value = "IPMT";
            worksheet["B23"].Formula = "=IPMT(0.1/12,1*3,3,8000)";
            worksheet["C23"].Formula = "=IPMT(0.1,3,3,8000)";

            worksheet["A24"].Value = "IRR";
            worksheet["B24"].Formula = "=IRR({-70000,12000,15000,18000,21000})";
            worksheet["C24"].Formula = "=IRR({-70000,12000,15000,18000,21000,26000})";
            worksheet["D24"].Formula = "=IRR({-70000,12000,15000},-0.1)";

            worksheet["A25"].Value = "ISPMT";
            worksheet["B25"].Formula = "=ISPMT(0.1/12,1,3*12,8000000)";
            worksheet["C25"].Formula = "=ISPMT(0.1,1,3,8000000)";
        }

        public static void Information(Workbook workbook)
        {
            var worksheet = workbook.AddWorksheet("Information");

            worksheet["A2"].Value = "ACCRINT";
            worksheet["B2"].Formula = "=ERROR.TYPE(F2)";
            worksheet["C2"].Formula = "=ERROR.TYPE(G2)";
            worksheet["D2"].Formula = "=ERROR.TYPE(H2)";
            worksheet["E2"].Formula = "=ERROR.TYPE(I2)";
            worksheet["F2"].Value = ErrorType.Div0;
            worksheet["G2"].Value = ErrorType.Ref;
            worksheet["H2"].Value = ErrorType.NA;
            worksheet["I2"].Value = "ABC";

            worksheet["A3"].Value = "ISBLANK";
            worksheet["B3"].Formula = "=ISBLANK(E3)";
            worksheet["C3"].Formula = "=ISBLANK(D3)";
            worksheet["D3"].Value = 2;

            worksheet["A4"].Value = "ISERR";
            worksheet["B4"].Formula = "=ISERR(D4)";
            worksheet["C4"].Formula = "=ISERR(E4)";
            worksheet["D4"].Value = ErrorType.Div0;
            worksheet["E4"].Value = ErrorType.NA;

            worksheet["A5"].Value = "ISERROR";
            worksheet["B5"].Formula = "=ISERROR(C5)";
            worksheet["C5"].Value = ErrorType.Div0;

            worksheet["A6"].Value = "ISEVEN";
            worksheet["B6"].Formula = "=ISEVEN(12.456)";
            worksheet["C6"].Formula = "=ISEVEN(D6)";
            worksheet["D6"].Value = -15;

            worksheet["A7"].Value = "ISLOGICAL";
            worksheet["B7"].Formula = "=ISLOGICAL(TRUE)";
            worksheet["C7"].Formula = "=ISLOGICAL(F7)";
            worksheet["D7"].Formula = "=ISLOGICAL({TRUE,2})";
            worksheet["E7"].Formula = "=ISLOGICAL({2,TRUE})";
            worksheet["F7"].Value = 123;

            worksheet["A8"].Value = "ISNA";
            worksheet["B8"].Formula = "=ISNA(D8)";
            worksheet["C8"].Formula = "=ISNA(E8)";
            worksheet["D8"].Value = ErrorType.Div0;
            worksheet["E8"].Value = ErrorType.NA;

            worksheet["A9"].Value = "ISNONTEXT";
            worksheet["B9"].Formula = "=ISNONTEXT(\"ABC\")";
            worksheet["C9"].Formula = "=ISNONTEXT(F9)";
            worksheet["D9"].Formula = "=ISNONTEXT({1,\"ABC\"})";
            worksheet["E9"].Formula = "=ISNONTEXT({\"ABC\",1})";
            worksheet["F9"].Value = 123;

            worksheet["A10"].Value = "ISNUMBER";
            worksheet["B10"].Formula = "=ISNUMBER(10.56)";
            worksheet["C10"].Formula = "=ISNUMBER(F10)";
            worksheet["D10"].Formula = "=ISNUMBER({1,\"ABC\"})";
            worksheet["E10"].Formula = "=ISNUMBER({\"ABC\",1})";
            worksheet["F10"].Value = "ABC";

            worksheet["A11"].Value = "ISODD";
            worksheet["B11"].Formula = "=ISODD(12.456)";
            worksheet["C11"].Formula = "=ISODD(D11)";
            worksheet["D11"].Value = -15;

            worksheet["A12"].Value = "ISREF";
            worksheet["B12"].Formula = "=ISREF(\"ABC\")";
            worksheet["C12"].Formula = "=ISREF(D12)";

            worksheet["A13"].Value = "ISTEXT";
            worksheet["B13"].Formula = "=ISTEXT(\"ABC\")";
            worksheet["C13"].Formula = "=ISTEXT(F13)";
            worksheet["D13"].Formula = "=ISTEXT({1,\"ABC\"})";
            worksheet["E13"].Formula = "=ISTEXT({\"ABC\",1})";
            worksheet["F13"].Value = 123;
        }

        public static void Logical(Workbook workbook)
        {
            var worksheet = workbook.AddWorksheet("Logical");

            worksheet["A1"].Value = "AND";
            worksheet["B1"].Formula = "=AND(TRUE)";
            worksheet["C1"].Formula = "=AND(TRUE,FALSE)";
            worksheet["D1"].Formula = "=AND(10>5,3=1+2,5)";

            worksheet["A2"].Value = "FALSE";
            worksheet["B2"].Formula = "=FALSE()";

            worksheet["A3"].Value = "IF";
            worksheet["B3"].Formula = "=IF(10>5,\"Yes\",\"No\")";
            worksheet["C3"].Formula = "=IF(10>5,\"Yes\")";
            worksheet["D3"].Formula = "=IF(10>5,\"Yes\",)";
            worksheet["E3"].Formula = "=IF(10<5,\"Yes\")";
            worksheet["F3"].Formula = "=IF(10<5,\"Yes\",)";

            worksheet["A4"].Value = "IFERROR";
            worksheet["B4"].Formula = "=IFERROR(1/0,\"Error in calculation\")";
        }

        public static void LookupReference(Workbook workbook)
        {
            var worksheet = workbook.AddWorksheet("LookupReference");

            worksheet["A1"].Value = "COLUMN";
            worksheet["B1"].Formula = "=COLUMN()";

            worksheet["A2"].Value = "COLUMNS";
            worksheet["B2"].Formula = "=COLUMNS(A5:C5)";
            worksheet["C2"].Formula = "=COLUMNS({1,2;3,4})";

            worksheet["A3"].Value = "INDEX";
            worksheet["B3"].Formula = "=INDEX({\"Apples\",\"Lemons\";\"Bananas\",\"Pears\"},2,2)";
            worksheet["C3"].Formula = "=INDEX({\"Apples\",\"Lemons\";\"Bananas\",\"Pears\"},2,1)";
            worksheet["D3"].Formula = "=INDEX({\"Apples\",\"Lemons\"},,2)";
            worksheet["E3"].Formula = "=INDEX({\"Apples\";\"Bananas\"},1)";
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

            worksheet["A9"].Value = "CEILING";
            worksheet["B9"].Formula = "=CEILING(2.5,1)";
            worksheet["C9"].Formula = "=CEILING(-2.5,-2)";
            worksheet["D9"].Formula = "=CEILING(1.5,0.1)";
            worksheet["E9"].Formula = "=CEILING(0.234,0.01)";

            worksheet["A10"].Value = "COMBIN";
            worksheet["B10"].Formula = "=COMBIN(8,2)";
            worksheet["C10"].Formula = "=COMBIN(10,4)";
            worksheet["D10"].Formula = "=COMBIN(6,5)";

            worksheet["A11"].Value = "COS";
            worksheet["B11"].Formula = "=COS(-1)";
            worksheet["C11"].Formula = "=COS(0)";
            worksheet["D11"].Formula = "=COS(1)";

            worksheet["A12"].Value = "COSH";
            worksheet["B12"].Formula = "=COSH(-1)";
            worksheet["C12"].Formula = "=COSH(0)";
            worksheet["D12"].Formula = "=COSH(1)";

            worksheet["A13"].Value = "DEGREES";
            worksheet["B13"].Formula = "=DEGREES(2 * PI())";
            worksheet["C13"].Formula = "=DEGREES(PI())";
            worksheet["D13"].Formula = "=DEGREES(PI()/2)";
            worksheet["E13"].Formula = "=DEGREES(8.5)";

            worksheet["A14"].Value = "EVEN";
            worksheet["B14"].Formula = "=EVEN(1.5)";
            worksheet["C14"].Formula = "=EVEN(3)";
            worksheet["D14"].Formula = "=EVEN(2)";
            worksheet["E14"].Formula = "=EVEN(-1)";

            worksheet["A15"].Value = "EXP";
            worksheet["B15"].Formula = "=EXP(-1)";
            worksheet["C15"].Formula = "=EXP(0)";
            worksheet["D15"].Formula = "=EXP(1)";
            worksheet["E15"].Formula = "=EXP(2)";

            worksheet["A16"].Value = "FACT";
            worksheet["B16"].Formula = "=FACT(5)";
            worksheet["C16"].Formula = "=FACT(3.5)";
            worksheet["D16"].Formula = "=FACT(0)";

            worksheet["A17"].Value = "FACTDOUBLE";
            worksheet["B17"].Formula = "=FACTDOUBLE(5)";
            worksheet["C17"].Formula = "=FACTDOUBLE(3.5)";
            worksheet["D17"].Formula = "=FACTDOUBLE(0)";

            worksheet["A18"].Value = "FLOOR";
            worksheet["B18"].Formula = "=FLOOR(2.5,1)";
            worksheet["C18"].Formula = "=FLOOR(-2.5,-2)";
            worksheet["D18"].Formula = "=FLOOR(1.5,0.1)";
            worksheet["E18"].Formula = "=FLOOR(0.234,0.01)";

            worksheet["A19"].Value = "GCD";
            worksheet["B19"].Formula = "=GCD(5)";
            worksheet["C19"].Formula = "=GCD(5,2)";
            worksheet["D19"].Formula = "=GCD(100,50,28)";
            worksheet["E19"].Formula = "=GCD(24.5,36.3)";
            worksheet["B19"].Formula = "=GCD(7,1)";
            worksheet["C19"].Formula = "=GCD(5,0)";

            worksheet["A20"].Value = "INT";
            worksheet["B20"].Formula = "=INT(8.9)";
            worksheet["C20"].Formula = "=INT(-8.9)";
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

            worksheet["A8"].Value = "CHIDIST";
            worksheet["B8"].Formula = "=CHIDIST(3.5,4)";
            worksheet["C8"].Formula = "=CHIDIST(12.34,7)";

            worksheet["A9"].Value = "CHIINV";
            worksheet["B9"].Formula = "=CHIINV(0.5,4)";
            worksheet["C9"].Formula = "=CHIINV(0.3,7)";

            worksheet["A9"].Value = "CONFIDENCE";
            worksheet["B9"].Formula = "=CONFIDENCE(0.4,5,12)";
            worksheet["C9"].Formula = "=CONFIDENCE(0.75,9,7)";

            worksheet["A10"].Value = "CORREL";
            worksheet["B10"].Formula = "=CORREL({2.532,5.621;2.1,3.4},{5.32,2.765;5.2,6.7})";

            worksheet["A11"].Value = "COUNTA";
            worksheet["B11"].Formula = "=COUNTA(1,2,3,4,5)";
            worksheet["C11"].Formula = "=COUNTA({1,2,3,4,5})";
            worksheet["D11"].Formula = "=COUNTA({1,2,3,4,5},6,\"7\")";
            worksheet["E11"].Formula = "=COUNTA(10,G11)";

            worksheet["A12"].Value = "COUNTBLANK";
            worksheet["D12"].Value = 2;
            worksheet["B12"].Formula = "=COUNTBLANK(C12:E12)";

            worksheet["A13"].Value = "COUNTIF Number";
            worksheet["E13"].Value = 3;
            worksheet["F13"].Value = 10;
            worksheet["F13"].Value = 7;
            worksheet["G13"].Value = 10;
            worksheet["B13"].Formula = "=COUNTIF(E13:G13,\"=10\")";
            worksheet["C13"].Formula = "=COUNTIF(E13:G13,\">5\")";
            worksheet["D13"].Formula = "=COUNTIF(E13:G13,\"<>10\")";


            worksheet["A14"].Value = "COUNTIF Text";
            worksheet["E14"].Value = "apples";
            worksheet["F14"].Value = "oranges";
            worksheet["F14"].Value = "grapes";
            worksheet["G14"].Value = "melons";
            worksheet["B14"].Formula = "=COUNTIF(E14:G14,\"*es\")";
            worksheet["C14"].Formula = "=COUNTIF(E14:G14,\"??a*\")";
            worksheet["D14"].Formula = "=COUNTIF(E14:G14,\"*l*\")";

            worksheet["A15"].Value = "COVAR";
            worksheet["B15"].Formula = "=COVAR({2.532,5.621;2.1,3.4},{5.32,2.765;5.2,6.7})";

            worksheet["A16"].Value = "CRITBINOM";
            worksheet["B16"].Formula = "=CRITBINOM(6,0.5,0.75)";
            worksheet["C16"].Formula = "=CRITBINOM(12,0.3,0.95)";

            worksheet["A17"].Value = "DEVSQ";
            worksheet["B17"].Formula = "=DEVSQ(5.6,8.2,9.2)";
            worksheet["C17"].Formula = "=DEVSQ({5.6,8.2,9.2})";

            worksheet["A18"].Value = "EXPONDIST";
            worksheet["B18"].Formula = "=EXPONDIST(0.2,10,FALSE)";
            worksheet["C18"].Formula = "=EXPONDIST(2.3,1.5,TRUE)";

            worksheet["A19"].Value = "FDIST";
            worksheet["B19"].Formula = "=FDIST(12.345,3,4)";

            worksheet["A20"].Value = "FINV";
            worksheet["B20"].Formula = "=FINV(0.5,3,4)";

            worksheet["A21"].Value = "FISHER";
            worksheet["B21"].Formula = "=FISHER(-0.43)";
            worksheet["C21"].Formula = "=FISHER(0.578)";

            worksheet["A22"].Value = "FISHERINV";
            worksheet["B22"].Formula = "=FISHERINV(-0.43)";
            worksheet["C22"].Formula = "=FISHERINV(0.578)";

            worksheet["A23"].Value = "FORECAST";
            worksheet["B23"].Formula = "=FORECAST(30,{6,7,9,15,21},{20,28,31,38,40})";

            worksheet["A24"].Value = "GAMMADIST";
            worksheet["B24"].Formula = "=GAMMADIST(10,9,2,FALSE)";
            worksheet["C24"].Formula = "=GAMMADIST(10,9,2,TRUE)";

            worksheet["A25"].Value = "GAMMAINV";
            worksheet["B25"].Formula = "=GAMMAINV(0.068,9,2)";

            worksheet["A26"].Value = "GAMMALN";
            worksheet["B26"].Formula = "=GAMMALN(4.5)";

            worksheet["A27"].Value = "GEOMEAN";
            worksheet["B27"].Formula = "=GEOMEAN(10.5,5.3,2.9)";
            worksheet["C27"].Formula = "=GEOMEAN(10.5,{5.3,2.9},\"12\")";

            worksheet["A28"].Value = "HARMEAN";
            worksheet["B28"].Formula = "=HARMEAN(4.6,5.8,8.3,7)";
            worksheet["C28"].Formula = "=HARMEAN(10.5,{5.3,2.9},\"12\")";

            worksheet["A29"].Value = "HYPGEOMDIST";
            worksheet["B29"].Formula = "=HYPGEOMDIST(1,4,8,20)";

            worksheet["A30"].Value = "INTERCEPT";
            worksheet["B30"].Formula = "=INTERCEPT({2,3,9,1,8},{6,5,11,7,5})";
        }

        public static void Text(Workbook workbook)
        {
            var worksheet = workbook.AddWorksheet("Text");

            worksheet["A1"].Value = "CHAR";
            worksheet["D1"].Value = 65;
            worksheet["B1"].Formula = "=CHAR(65)";
            worksheet["C1"].Formula = "=CHAR(D1)";

            worksheet["A2"].Value = "CLEAN";
            worksheet["B2"].Formula = "=CLEAN(\"A\" & CHAR(2) & \"BC\")";

            worksheet["A3"].Value = "CODE";
            worksheet["B3"].Formula = "=CODE(\"abc\")";

            worksheet["A4"].Value = "CONCATENATE";
            worksheet["B4"].Formula = "=CONCATENATE(3,\" + \",4,\" = \",3+4)";

            worksheet["A5"].Value = "DOLLAR";
            worksheet["B5"].Formula = "=DOLLAR(1234.567)";
            worksheet["C5"].Formula = "=DOLLAR(1234.567,-2)";
            worksheet["D5"].Formula = "=DOLLAR(-1234.567,4)";

            worksheet["A6"].Value = "EXACT";
            worksheet["B6"].Formula = "=EXACT(\"ABC\",\"ABC\")";
            worksheet["C6"].Formula = "=EXACT(\"ABC\",\"ABCD\")";
            worksheet["D6"].Formula = "=EXACT(\"Abc\",\"aBC\")";
            worksheet["E6"].Formula = "=EXACT(\"\",\"\")";

            worksheet["A7"].Value = "FIND";
            worksheet["B7"].Formula = "=FIND(\"de\",\"abcdef\")";

            worksheet["A8"].Value = "FIXED";
            worksheet["B8"].Formula = "=FIXED(1234567)";
            worksheet["C8"].Formula = "=FIXED(1234567.555555,4,TRUE)";
            worksheet["D8"].Formula = "=FIXED(.555555,10)";
            worksheet["E8"].Formula = "=FIXED(1234567,-3)";
        }
    }
}
