using MajConsoSuiviSprint.Cli.Model;
using MajConsoSuiviSprint.Cli.Helper;
using MajConsoSuiviSprint.Cli.Utils;
using MajConsoSuiviSprint.Cli.Constants;

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
        private bool ErrorSaisie(string activite, string numDemande, string nomAppli, WebTTTInfoModel infoWebTTT)
        {
            bool result = true;


            if (activite == infoWebTTT.MaskSaisieAutorise["Spec"].ToString())
            {
                infoWebTTT.MaskSaisieAutorise["Spec"].ForEach(mask => 
                {
                    if (activite.Contains(mask.Contain))
                    {
                        result = false;
                        break;
                    }
                });
                if(result)
                {
                    result = (nomAppli == AppliConstant.LblApplicationParDefaut);
                }
            }

            else if (activite == infoWebTTT.MaskSaisieAutorise["DevQual"].ToString())
            {
                infoWebTTT.MaskSaisieAutorise["DevQual"].ForEach(mask =>
                {
                    if (activite.Contains(mask.Contain))
                    {
                        result = false;
                    }
                });

            }
               
            return result;
        }

        private bool SaisieAprendreEnCompte(string activite,DateTime dateSaisie, WebTTTInfoModel infoWebTTT)
        {
            int numDebut=  (infoWebTTT.NumDebutSemaineAImporter - (infoWebTTT.NbreSprintAPrendreEnCompte * 2)) ;
            return (InfoSprint.IsActivityToManaged(activite) && InfoSprint.IsPeriodeToManaged(dateSaisie, numDebut, infoWebTTT.NumFinSemaineAImporter));
        }
    }
}
