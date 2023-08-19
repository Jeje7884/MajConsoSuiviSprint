using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace MajConsoSuiviSprint.Cli.Helper
{
    internal class ExceLNPOIHelper : IDisposable
    {
        public IWorkbook Workbook;


        private bool Disposed = false;


        ~ExceLNPOIHelper() => Dispose(false);

        // Public implementation of Dispose pattern callable by consumers.
        public void Dispose()
        {
            // Dispose of unmanaged resources.
            Dispose(true);
            // Suppress finalization.
            GC.SuppressFinalize(this);
        }

        // Protected implementation of Dispose pattern.
        protected virtual void Dispose(bool disposing)
        {
            if (Disposed)
            {
                return;
            }

            if (disposing)
            {
                Workbook.Dispose();
            }

            Disposed = true;
        }

        public ExceLNPOIHelper(string pathFile)
        {
            FileStream fileStream = new(pathFile, FileMode.Open, FileAccess.Read);
            Workbook = new XSSFWorkbook(fileStream);

        }
        public ISheet GetSheet(string sheetName)
        {
            return Workbook.GetSheet(sheetName);
        }

     

        public static Dictionary<string, int> GetIdColumns(IRow headerRow, List<string> columnsToImport)
        {
            Dictionary<string, int> dictionnaireColonne = new();

            if (columnsToImport?.Count > 0)
            {
                foreach (string column in columnsToImport)
                {
                    int numColonne = GetColumnIndex(headerRow, column);

                    dictionnaireColonne.Add(column, numColonne);
                }
            }

            return dictionnaireColonne;
        }



        private static int GetColumnIndex(IRow headerRow, string columnName)
        {
            for (int i = 0; i < headerRow.Cells.Count; i++)
            {
                if (headerRow.Cells[i].StringCellValue.Equals(columnName, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }
            return -1;
        }

        public static string GetCellValue(IRow row, int columnIndex)
        {
            if (columnIndex >= 0 && columnIndex < row.Cells.Count)
            {
                return row.Cells[columnIndex].ToString() ?? "";
            }
            return string.Empty;
        }


    }
}