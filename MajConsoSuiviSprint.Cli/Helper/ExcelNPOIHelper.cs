using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;


namespace MajConsoSuiviSprint.Cli.Helper
{
    internal static class ExceLNPOIHelper
    {
        public static  List<Dictionary<string,string>> ImportExcel(string pathFile, string sheetName, List<string> columnsToImport)
        {

            using FileStream fs = new(pathFile, FileMode.Open, FileAccess.Read);
            IWorkbook workbook = new XSSFWorkbook(fs);
            ISheet sheet = workbook.GetSheet(sheetName); // Indice de la feuille (0 pour la première feuille)

            var header = sheet.GetRow(0);
            var columnIndexToHeader = Enumerable.Range(0, header.Cells.Count)
                .ToDictionary(index => index, index => header.GetCell(index).StringCellValue);

            var importedData = Enumerable.Range(1, sheet.LastRowNum)
                .Select(rowIndex =>
                {
                    var dataRow = sheet.GetRow(rowIndex);
                    return columnIndexToHeader
                        .Where(kv => columnsToImport.Contains(kv.Value))
                        .ToDictionary(kv => kv.Value, kv => dataRow.GetCell(kv.Key)?.ToString() ?? "");
                })
                .ToList();
            return importedData;

        }
    }
}
