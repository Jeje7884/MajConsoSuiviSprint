using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace MajConsoSuiviSprint.Cli.Helper
{
    internal class EPPlus
    {
        internal static void GetInfoFichierSuiviEEPlus()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using ExcelPackage package = new(new FileInfo(@"C:\temp\cd13\AppliSuiviWebTTT\CD13_PI08_S29-30-testEpplus-2.xlsx"));

            ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Sélectionnez la feuille

            ExcelTable table = worksheet.Tables["TableauSuiviSprint"]; // Sélectionnez la table

            // Accédez aux données de la table
            foreach (var row in table.WorkSheet.Cells[table.Address.FullAddress].Skip(26))
            {
                var val = worksheet.Cells[row.Address].Text;
                //var test = row.Offset(1, 3).First();
                if (row.Address.Contains("AH"))
                {
                    int numligne = row.EntireRow.StartRow;
                    Console.WriteLine($"{row.Address} N° de demande   {val} ");
      
                    var startCell = worksheet.Cells[$"AR{numligne}"];
                    startCell.Value = 10;

                }


            }
            // worksheet.Calculate();               

            //package.Save(); ne marche pas. Verrole le fichier
        }

        public static void ReadXLS()
        {
            FileInfo existingFile = new FileInfo(@"C:\temp\cd13\AppliSuiviWebTTT\CD13_PI08_S29-30-testEpplus.xlsx");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count
                for (int row = 1; row <= rowCount; row++)
                {
                    for (int col = 1; col <= colCount; col++)
                    {
                        Console.WriteLine(" Row:" + row + " column:" + col + " Value:" + worksheet?.Cells?[row, col]?.Value?.ToString()?.Trim());
                    }
                }
                package.Save();
            }
        }
    }
}
