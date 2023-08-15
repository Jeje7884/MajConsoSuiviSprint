using MajConsoSuiviSprint.Cli.Business.Interfaces;
using MajConsoSuiviSprint.Cli.Constants;
using MajConsoSuiviSprint.Cli.Helper;
using MajConsoSuiviSprint.Cli.Model;
using MajConsoSuiviSprint.Cli.Utils;

namespace MajConsoSuiviSprint.Cli.Business
{
    internal class ImportWebTTT
    {

        public List<ImportWebTTTExcelModel> ImportSaisisWebTTTExcel = new();
        public List<ResultatImport> ResultatsWebTTTExcel = new();
        public List<ErreurSaisieDemandeModel> ErreursDeSaisies = new();
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
            public List<ErreurSaisieDemandeModel> ErrorsSaisis { get; set; } = new();
            public Dictionary<string, TimeConsumedByActivite> TempsConsommesParDemandes { get; set; } = new();

        }

        internal ResultatImport ImportInfosFromWebTTT()
        {

            if (Divers.IsFileOpened(_configurationApp.WebTTTInfoConfig.FullFileName))
            {
                throw new ExceptFileOpenException($"Le fichier {_configurationApp.WebTTTInfoConfig.FullFileName} est ouvert");
            }
            var result = ExceLNPOIHelper.ImportFichierWebTTTExcel(_configurationApp.WebTTTInfoConfig);
            Console.WriteLine($" {result.Count} lignes ont été récupérés du fichier {_configurationApp.WebTTTInfoConfig.FileName}");

            return new ResultatImport();
        }
        private bool ErrorSaisie(string activite, string numDemande, string nomAppli, WebTTTInfoConfigModel infoWebTTT)
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

        private bool SaisieAprendreEnCompte(string activite, DateTime dateSaisie, WebTTTInfoConfigModel infoWebTTT)
        {
            return (InfoSprint.IsActivityToManaged(activite) && InfoSprint.IsPeriodeToManaged(dateSaisie, infoWebTTT.NumeroDebutSemaineAImporter, infoWebTTT.NumeroFinSemaineAImporter));
        }
    }
}

