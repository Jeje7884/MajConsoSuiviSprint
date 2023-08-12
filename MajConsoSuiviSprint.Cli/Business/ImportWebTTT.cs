using MajConsoSuiviSprint.Cli.Model;
using MajConsoSuiviSprint.Cli.Helper;

namespace MajConsoSuiviSprint.Cli.Business
{
    internal class ImportWebTTT
    {

        //public Dictionary<string, TimeConsumedByActivite> TempsConsommesParDemandes  = new();
        //public List<ErrorSaisieWebTTTModel> ErrorsSaisisWebTTT = new();
        public List<ImportWebTTTExcelModel> ImportSaisisWebTTTExcel = new();
        public List<ResultatImport> ResultatsWebTTTExcel = new();

        public class TimeConsumedByActivite
        {
            public int HourQual { get; set; } = 0;

            public int HourDev { get; set; } = 0;
            public string Application { get; set; } = string.Empty;
            public bool NumDemandeIsValid { get; set; } = true;
        }

        public class ResultatImport
        {
            public List<ErrorSaisieWebTTTModel> ErrorsSaisis { get; set; } = new();
            public Dictionary<string, TimeConsumedByActivite> TempsConsommesParDemandes { get; set; } = new();

        }
        public ImportWebTTT() { 
        }

        public ResultatImport ImportInfosFromWebTTT(string path, string sheetName, List<string> columnsToImport)
        {
            var result = ExceLNPOIHelper.ImportExcel(path, sheetName, columnsToImport);
            return new ResultatImport();
        }
    }
}
