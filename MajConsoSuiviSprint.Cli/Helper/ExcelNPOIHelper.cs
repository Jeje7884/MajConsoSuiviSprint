using MajConsoSuiviSprint.Cli.Model;
using MajConsoSuiviSprint.Cli.Utils;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace MajConsoSuiviSprint.Cli.Helper
{
    internal static class ExceLNPOIHelper
    {
        public static List<ImportWebTTTExcelModel> ImportFichierWebTTTExcel(WebTTTInfoConfigModel WebTTTInfoConfigModel,List<string> columnsToImport)

        {
            string pathFile = WebTTTInfoConfigModel.FullFileName;

            string sheetName = WebTTTInfoConfigModel.SheetName;
     

            var result = new List<ImportWebTTTExcelModel>();
            var dictionnaireNumeroDecolonne = new Dictionary<string, int>();

            using (FileStream fileStream = new(pathFile, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(fileStream);
                ISheet worksheet = workbook.GetSheet(sheetName);

                IRow headerRow = worksheet.GetRow(0);

                dictionnaireNumeroDecolonne = InitDictionnaireNumeroColonne(headerRow, columnsToImport);
                Dictionary<int, string> columnIndexToHeader = new();

                for (int row = 1; row <= worksheet.LastRowNum; row++)
                {
                    IRow dataRow = worksheet.GetRow(row);

                    if (dataRow != null)
                    {
                        var data = GetDataFromRowWebTTT(WebTTTInfoConfigModel, result, dictionnaireNumeroDecolonne, dataRow);
                        if (null != data)
                        {
                            result.Add(data);
                        }
                    }
                }
            }
            return result;
        }

        private static ImportWebTTTExcelModel GetDataFromRowWebTTT(WebTTTInfoConfigModel WebTTTInfoConfigModel, List<ImportWebTTTExcelModel> result, Dictionary<string, int> dictionnaireNumeroDecolonne, IRow dataRow)
        {
            DateTime dateDeSaisie = DateTime.Parse(GetCellValue(dataRow, dictionnaireNumeroDecolonne["Date"]));           
            int numeroDeSemaineDateActivite = InfoSprint.GetNumSemaine(dateDeSaisie);
            ImportWebTTTExcelModel dataFromRowWebTTT = default!;

            if (numeroDeSemaineDateActivite >= WebTTTInfoConfigModel.NumeroDeSemaineAPartirDuquelChecker)
            {
                dataFromRowWebTTT = new ImportWebTTTExcelModel
                {
                    TrigrammeCollab = GetCellValue(dataRow, dictionnaireNumeroDecolonne["Collaborator"]) ?? default!,
                    Activite = GetCellValue(dataRow, dictionnaireNumeroDecolonne["Activity"]) ?? default!,
                    Application = GetCellValue(dataRow, dictionnaireNumeroDecolonne["Domain"]) ?? default!,
                    DateDeSaisie = dateDeSaisie,
                    NumeroDeDemande = GetCellValue(dataRow, dictionnaireNumeroDecolonne["Demand number"]) ?? default!,
                    HeureDeclaree = float.Parse(GetCellValue(dataRow, dictionnaireNumeroDecolonne["Hours"])),
                    NumeroDeSemaineDateActivite = numeroDeSemaineDateActivite
                };
            }
            return dataFromRowWebTTT;
        }

        private static Dictionary<string, int> InitDictionnaireNumeroColonne(IRow headerRow, List<string> columnsToImport)
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

        private static string GetCellValue(IRow row, int columnIndex)
        {
            if (columnIndex >= 0 && columnIndex < row.Cells.Count)
            {
                return row.Cells[columnIndex].ToString() ?? "";
            }
            return string.Empty;
        }
    }
}