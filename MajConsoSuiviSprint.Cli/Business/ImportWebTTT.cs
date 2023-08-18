using MajConsoSuiviSprint.Cli.Business.Interfaces;
using MajConsoSuiviSprint.Cli.Constants;
using MajConsoSuiviSprint.Cli.Helper;
using MajConsoSuiviSprint.Cli.Model;
using MajConsoSuiviSprint.Cli.Utils;

namespace MajConsoSuiviSprint.Cli.Business
{
    internal class ImportWebTTT : IImportWebTTT
    {
        private readonly IConfigurationsApp _configurationApp;

        public ImportWebTTT(IConfigurationsApp configuration)
        {
            _configurationApp = configuration;
        }

        internal IList<ImportWebTTTExcelModel> ImportInfosFromWebTTT()
        {
            if (Divers.IsFileOpened(_configurationApp.WebTTTInfoConfig.FullFileName))
            {
                throw new ExceptFileOpenException($"Le fichier {_configurationApp.WebTTTInfoConfig.FullFileName} est ouvert");
            }
            var result = ExceLNPOIHelper.ImportFichierWebTTTExcel(_configurationApp.WebTTTInfoConfig);
            Console.WriteLine($" {result.Count} lignes ont été récupérés du fichier {_configurationApp.WebTTTInfoConfig.FileName}");

            return result;

        }

        public ResultImportWebTTT CheckSaisiesActiviteInWebTTT(IList<ImportWebTTTExcelModel> allDataInWebTTT)
        {

            List<ErreurSaisieDemandeModel> erreursSaisiesDemandes = new();

            List<SaisieRemplissageTempsCollabParSemaineModel> saisiesRemplissageTempsCollabParSemaine = new();

            foreach (ImportWebTTTExcelModel dataRowWebTTT in allDataInWebTTT)
            {
                bool isDemandeValid = InfoSprint.IsDemandeValid(dataRowWebTTT, _configurationApp.WebTTTInfoConfig.ReglesSaisiesAutorisesParActivite);

                if (!isDemandeValid)
                {
                    AddListErreursDeSaisies(erreursSaisiesDemandes, dataRowWebTTT);
                }
                AddListTempsDeclareParCollabEtParSemaine(saisiesRemplissageTempsCollabParSemaine, dataRowWebTTT);
            
            }
            return new ResultImportWebTTT()
            {
                ErreursSaisiesDemandes = erreursSaisiesDemandes,
                SaisiesRemplissageTempsCollabParSemaine = saisiesRemplissageTempsCollabParSemaine,
            
            };
        }

        public Dictionary<string, TempsConsommeDemandeModel>  GetTempsConsommeesParDemandeSurUneSemaine(IList<ImportWebTTTExcelModel> saisiesActivitesInWebTTT)
        {
            var result = new Dictionary<string, TempsConsommeDemandeModel>();
            var saisiesActivitesSurLeSprint = saisiesActivitesInWebTTT.Where(data => data.NumeroDeSemaineDateActivite.Equals(_configurationApp.SuiviSprintInfoConfig.NumeroSemaineDebutDeSprint) ||
             data.NumeroDeSemaineDateActivite.Equals(_configurationApp.SuiviSprintInfoConfig.NumeroSemaineFinDeSprint));

            foreach (var saisieActivite in saisiesActivitesSurLeSprint)
            {

                string numeroFormatSuiviSprint = InfoSprint.ModifyDemandeWebTTTToSuiviSprint(saisieActivite.NumeroDeDemande,saisieActivite.Activite);
                if (InfoSprint.IsActivityToManaged(saisieActivite.Activite))
                {
                    if (result.ContainsKey(numeroFormatSuiviSprint))
                    {
                        if (saisieActivite.Activite.Equals(AppliConstant.LblActiviteQual))
                        {
                            result[numeroFormatSuiviSprint].HeureTotaleDeQualification += saisieActivite.HeureDeclaree;
                        }
                        else
                        {
                            result[numeroFormatSuiviSprint].HeureTotaleDeDeveloppement += saisieActivite.HeureDeclaree;
                        }
                        
                    }
                    else
                    {
                        var conso = new TempsConsommeDemandeModel
                        {
                            Application = saisieActivite.Application,
                            HeureTotaleDeDeveloppement = saisieActivite.Activite.Equals(AppliConstant.LblActiviteDev) ? saisieActivite.HeureDeclaree : 0,
                            HeureTotaleDeQualification = saisieActivite.Activite.Equals(AppliConstant.LblActiviteQual) ? saisieActivite.HeureDeclaree : 0,
                            IsDemandeValide = InfoSprint.IsDemandeValid(saisieActivite,_configurationApp.WebTTTInfoConfig.ReglesSaisiesAutorisesParActivite)

                        };
                        result.Add(numeroFormatSuiviSprint, conso);
                    }
                    
                }
               
            }

            return result;
        }

     

        private static void AddListTempsDeclareParCollabEtParSemaine(IList<SaisieRemplissageTempsCollabParSemaineModel> tempsDeclaresParCollab, ImportWebTTTExcelModel dataRowWebTTT)
        {

            if (!tempsDeclaresParCollab.Any(tps => tps.Qui.Equals(dataRowWebTTT.TrigrammeCollab)
                                                                                        && tps.NumeroDeSemaine.Equals(dataRowWebTTT.NumeroDeSemaineDateActivite)))
            {
                SaisieRemplissageTempsCollabParSemaineModel saisieDemande = new()
                {
                    Qui = dataRowWebTTT.TrigrammeCollab,
                    NumeroDeSemaine = dataRowWebTTT.NumeroDeSemaineDateActivite,                    
                    TotalHeureDeclaree = dataRowWebTTT.HeureDeclaree
                };
                tempsDeclaresParCollab.Add(saisieDemande);
            }
            else
            {
                var tempsDuCollabEtPourUneSemaine = tempsDeclaresParCollab
                                                                                        .FirstOrDefault(tps =>
                                                                                            tps.Qui.Equals(dataRowWebTTT.TrigrammeCollab)
                                                                                            && tps.NumeroDeSemaine.Equals(dataRowWebTTT.NumeroDeSemaineDateActivite)
                                                                                            )
                                                                                            ?? default!;


                tempsDuCollabEtPourUneSemaine.TotalHeureDeclaree += dataRowWebTTT.HeureDeclaree;

            }
        }


        private static void AddListErreursDeSaisies(IList<ErreurSaisieDemandeModel> erreursSaisiesDemandes, ImportWebTTTExcelModel dataWebTTT)
        {
            if (!erreursSaisiesDemandes.Any(err => err.Equals(dataWebTTT.NumeroDeDemande)))
            {
                ErreurSaisieDemandeModel erreurSaisieDemande = new()
                {
                    Activite = dataWebTTT.Activite,
                    Application = dataWebTTT.Application,
                    NumeroDeDemande = dataWebTTT.NumeroDeDemande,
                    Qui = dataWebTTT.TrigrammeCollab,
                    DateDeSaisie = dataWebTTT.LblDateDeSaisie,
                    DetailErreur = AppliConstant.MessageErreurSaisie
                };
                erreursSaisiesDemandes.Add(erreurSaisieDemande);
            }
        }

       

        public void GenereExportCSVErreurSaisies(IList<ErreurSaisieDemandeModel> erreursSaisiesDemandes)
        {
            CSVHelper.GenerateCSVFile(_configurationApp.WebTTTInfoConfig.FileBilansErreurFormatSaisieDemande, erreursSaisiesDemandes, false);
        }

        public void GenereExportCSVErreurTempsConsomme(IList<SaisieRemplissageTempsCollabParSemaineModel> tempsConsommes)
        {
            var erreurDeTEmps = tempsConsommes.Where(tpsCollabParSemaine =>
                                                                            tpsCollabParSemaine.TotalHeureDeclaree < _configurationApp.WebTTTInfoConfig.NbreHeureTotaleMinimumAdeclarerParCollabEtParSemaine);


            CSVHelper.GenerateCSVFile(_configurationApp.WebTTTInfoConfig.FileBilansErreurTempsSaisieSemaine, tempsConsommes, false);
        }

       
    }
}