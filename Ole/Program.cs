using DRIT.Spreadsheet;

namespace Ole
{
    class Program
    {
        static void Main(string[] args)
        {
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];

            worksheet.Columns["B"].WidthPixels = 250;
            worksheet.GetRange("A1:A10").SetRowsHeight(53);

            var ole = worksheet.OleObjects.Embed("A1", @"..\In\ToEmbed.xlsx", @"..\In\OleIcon.emf");
            ole.Size.HeightInches = 0.75;
            ole.Size.WidthInches = 1;

            workbook.SaveAs(@"..\Out\OleEmbedIcon.xlsx");
        }
    }
}
