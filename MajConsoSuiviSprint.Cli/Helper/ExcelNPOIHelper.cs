using MajConsoSuiviSprint.Cli.Model;
using MajConsoSuiviSprint.Cli.Utils;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace MajConsoSuiviSprint.Cli.Helper
{
    internal static class ExceLNPOIHelper
    {
        public static List<ImportWebTTTExcelModel> ImportFichierWebTTTExcel(WebTTTInfoConfigModel WebTTTInfoConfigModel)

        {
            string pathFile = WebTTTInfoConfigModel.FullFileName;

            string sheetName = WebTTTInfoConfigModel.SheetName;
            List<HeadersWebTTTModel> columnsToImport = WebTTTInfoConfigModel.Headers.ToList();

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
                        var data= GetDataFromRowWebTTT(WebTTTInfoConfigModel, result, dictionnaireNumeroDecolonne, dataRow);
                        if (null!=data)
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
            string collab = GetCellValue(dataRow, dictionnaireNumeroDecolonne["Collaborator"]) ?? "";
            DateTime dateDeSaisie = DateTime.Parse(GetCellValue(dataRow, dictionnaireNumeroDecolonne["Date"]));
            string activite = GetCellValue(dataRow, dictionnaireNumeroDecolonne["Activity"]) ?? "";
            string application = GetCellValue(dataRow, dictionnaireNumeroDecolonne["Domain"]) ?? "";
            float heureDeclaree = float.Parse(GetCellValue(dataRow, dictionnaireNumeroDecolonne["Hours"]));
            int numeroDeSemaineDateActivite = InfoSprint.GetNumSemaine(dateDeSaisie);
            ImportWebTTTExcelModel dataFromRowWebTTT=default!;
            //if (IsSaisieAPrendreEnCompte(numeroDeSemaineDateActivite, WebTTTInfoConfigModel))
            if (numeroDeSemaineDateActivite >= WebTTTInfoConfigModel.NumeroDeSemaineAPartirDuquelChecker)
            {
                dataFromRowWebTTT = new ImportWebTTTExcelModel
                {
                    TrigrammeCollab = collab,
                    Activite = activite,
                    Application = application,
                    DateDeSaisie = dateDeSaisie,
                    HeureDeclaree = heureDeclaree,
                    NumeroDeSemaineDateActivite = numeroDeSemaineDateActivite
                } ;
            }
            return dataFromRowWebTTT;
        }

        //private static bool IsSaisieAPrendreEnCompte(int numeroDeSemaineDateActivite, WebTTTInfoConfigModel webTTTInfoConfigModel)
        //{
        //    int numDebutAPrendreEncompte = (webTTTInfoConfigModel.NumeroDebutSemaineAImporter - (webTTTInfoConfigModel.NbreSprintAPrendreEnCompte * 2));
        //    return numeroDeSemaineDateActivite >= numDebutAPrendreEncompte && numDebutAPrendreEncompte <= webTTTInfoConfigModel.NumeroFinSemaineAImporter;

        //}

        private static Dictionary<string, int> InitDictionnaireNumeroColonne(IRow headerRow, List<HeadersWebTTTModel> columnsToImport)
        {
            Dictionary<string, int> dictionnaireColonne = new();

            if (columnsToImport?.Count > 0)
            {
                foreach (HeadersWebTTTModel column in columnsToImport)
                {
                    int numColonne = GetColumnIndex(headerRow, column.Value);

                    dictionnaireColonne.Add(column.Value, numColonne);
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