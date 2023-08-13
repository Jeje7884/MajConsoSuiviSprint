using MajConsoSuiviSprint.Cli.Business.Interfaces;
using MajConsoSuiviSprint.Cli.Constants;
using MajConsoSuiviSprint.Cli.Helper;
using MajConsoSuiviSprint.Cli.Model;
using MajConsoSuiviSprint.Cli.Utils;
using Microsoft.Extensions.Configuration;

namespace MajConsoSuiviSprint.Cli.Business
{
    internal class ImportWebTTT
    {

        public List<ImportWebTTTExcelModel> ImportSaisisWebTTTExcel = new();
        public List<ResultatImport> ResultatsWebTTTExcel = new();
        public List<ErrorSaisieWebTTTModel> ErreursDeSaisies = new();
        private readonly IConfigurationsApp _configurationApp;

        public ImportWebTTT(IConfigurationsApp configuration)
        {
            _configurationApp = configuration;
        }

        internal class TimeConsumedByActivite
        {
            public int HourQual { get; set; } = 0;

            public int HourDev { get; set; } = 0;
            public string Application { get; set; } = string.Empty;
            public bool NumDemandeIsValid { get; set; } = true;
        }

        internal class ResultatImport
        {
            public List<ErrorSaisieWebTTTModel> ErrorsSaisis { get; set; } = new();
            public Dictionary<string, TimeConsumedByActivite> TempsConsommesParDemandes { get; set; } = new();

        }


        internal ResultatImport ImportInfosFromWebTTT(string path, string sheetName, List<string> columnsToImport)
        {
            var result = ExceLNPOIHelper.ImportExcel(path, sheetName, columnsToImport);
            return new ResultatImport();
        }
        private bool ErrorSaisie(string activite, string numDemande, string nomAppli, WebTTTInfoModel infoWebTTT)
        {
            bool result = true;


            if (activite == infoWebTTT.ReglesSaisiesAutorisesParActivite["Spec"].ToString())
            {
                infoWebTTT.ReglesSaisiesAutorisesParActivite["Spec"].ForEach(mask =>
                {
                    if (activite.Contains(mask.Rule))
                    {
                        result = false;

                    }
                });
                if (result)
                {
                    result = (nomAppli == AppliConstant.LblApplicationParDefaut);
                }
            }

            else if (activite == infoWebTTT.ReglesSaisiesAutorisesParActivite["DevQual"].ToString())
            {
                infoWebTTT.ReglesSaisiesAutorisesParActivite["DevQual"].ForEach(mask =>
                {
                    if (activite.Contains(mask.Rule))
                    {
                        result = false;
                    }
                });

            }

            return result;
        }

        private bool SaisieAprendreEnCompte(string activite, DateTime dateSaisie, WebTTTInfoModel infoWebTTT)
        {
            int numDebut = (infoWebTTT.NumDebutSemaineAImporter - (infoWebTTT.NbreSprintAPrendreEnCompte * 2));
            return (InfoSprint.IsActivityToManaged(activite) && InfoSprint.IsPeriodeToManaged(dateSaisie, numDebut, infoWebTTT.NumFinSemaineAImporter));
        }
    }
}

