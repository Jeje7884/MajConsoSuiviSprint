using ClosedXML.Excel;


namespace MajConsoSuiviSprint.Cli.Helper
{
    internal class ClosedXmlHelper : IDisposable
    {
        public XLWorkbook Workbook;
        private bool Disposed = false;


        ~ClosedXmlHelper() => Dispose(false);

        public ClosedXmlHelper(string filePath)
        {
            Workbook = new XLWorkbook(filePath);
        }

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

        /// Retourne la colonne correspondante au chiffre passé
        /// </summary>
        /// <param name="column">Numéro de la colonne</param>
        public static string ExcelColumnFromInt(int column)
        {
            string columnString = "";
            decimal columnNumber = column;
            while (columnNumber > 0)
            {
                decimal currentLetterNumber = (columnNumber - 1) % 26;
                char currentLetter = (char)(currentLetterNumber + 65);
                columnString = currentLetter + columnString;
                columnNumber = (columnNumber - (currentLetterNumber + 1)) / 26;
            }
            return columnString;
        }

        /// <summary>
        /// Recupere une workSheet
        /// </summary>
        public IXLWorksheet GetWorksheetByName(string sheetName)
        {
            IXLWorksheet result = default!;

            if (Workbook.TryGetWorksheet(sheetName, out IXLWorksheet worksheet))
            {
                result = worksheet;
            }

            return result;
        }

        public IXLWorksheet GetWorksheetById(int idSheet)
        {

            IXLWorksheet sheet = default!;

            if (idSheet >= 0 && idSheet < Workbook.Worksheets.Count)
            {
                sheet = Workbook.Worksheets.Worksheet(idSheet);

                // Maintenant, vous pouvez travailler avec la feuille par son index
            }
            else
            {
                Console.WriteLine("Index de feuille invalide.");
            }


            return sheet;
        }

       

        /// <summary>
        /// Recupere le premier tableau d'une worksheet
        /// </summary>
        public static IXLTable GetFirstTable(IXLWorksheet worksheet)
        {
            return worksheet.Tables.FirstOrDefault() ?? default!;
        }

        public static IXLTable GetTableByName(IXLWorksheet worksheet, string tableName)
        {
            return worksheet.Tables.Table(tableName);
        }

        public static void SetValueInCell(IXLCell cell, float value)
        {
            cell.Value = value;
        }

        public static string GetValueInCellFromTable(IXLTableRow row, int idColumn)
        {
            return row.Cell(idColumn).GetString();
        }

        public static IXLRangeColumn GetFistColumnInTable(IXLTable table)
        {
            return table.FirstColumnUsed();
        }

        public static void StrickeCell(IXLTableRow row, int columnIndex)
        {

            row.Cell(columnIndex).Style.Font.Strikethrough = true; 

        }


        public static void SetValueCell(IXLTableRow row , int columnIndex, string value)
        {

            row.Cell(columnIndex).Value = value;


        }
    }

}
