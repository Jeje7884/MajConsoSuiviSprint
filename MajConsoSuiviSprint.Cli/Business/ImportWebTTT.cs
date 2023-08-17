using DocumentFormat.OpenXml.Office2013.Word;
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

            //ResultImportWebTTT resultImport = CheckSaisiesActiviteInWebTTT(result);


            //return new ResultImportWebTTT();
            return result;

        }

        public ResultImportWebTTT CheckSaisiesActiviteInWebTTT(IList<ImportWebTTTExcelModel> allDataInWebTTT)
        {
            //Dictionary<string, ErreurSaisieDemandeModel> erreursSaisiesDemandes = new();
            List<ErreurSaisieDemandeModel> erreursSaisiesDemandes = new();


            //Dictionary<string, TempsConsommeDemandeModel> tempsConsommesPardemandesEtParSprint = new();

            //Dictionary<string, Dictionary<int, float>> tempsDeclaresParCollab = new();
            List<SaisieRemplissageTempsCollabParSemaineModel> saisiesRemplissageTempsCollabParSemaine = new();

            foreach (ImportWebTTTExcelModel dataRowWebTTT in allDataInWebTTT)
            {
                bool isDemandeValid = IsSaisieValid(dataRowWebTTT);

                if (!isDemandeValid)
                {
                    AddListErreursDeSaisies(erreursSaisiesDemandes, dataRowWebTTT);
                }
                //if (SaisieAPrendreEnComptePourTempsConsoFichierDeSuivi(dataWebTTT))
                //{
                //    AddDictionnaireTempsConsommeParDemandeEtParPole(tempsConsommesPardemandesEtParSprint, dataWebTTT);

                //    tempsConsommesPardemandesEtParSprint[dataWebTTT.NumeroDeDemande].IsDemandeValide = isDemandeValid;
                //}
                AddListTempsDeclareParCollabEtParSemaine(saisiesRemplissageTempsCollabParSemaine, dataRowWebTTT);
                //erreursSaisiesRemplissageTempsCollabParSemaine = GetListCollabErreurRemplissageHeruresWebTTTT(tempsDeclaresParCollab);
            }
            return new ResultImportWebTTT()
            {
                ErreursSaisiesDemandes = erreursSaisiesDemandes,
                SaisiesRemplissageTempsCollabParSemaine = saisiesRemplissageTempsCollabParSemaine,
                //TempsConsommesPardemandesEtParSprint = tempsConsommesPardemandesEtParSprint
            };
        }

        //private List<SaisieRemplissageTempsCollabParSemaineModel> GetListCollabErreurRemplissageHeruresWebTTTT(Dictionary<string, Dictionary<int, float>> listSaisiesTotale)
        //{
        //    List<SaisieRemplissageTempsCollabParSemaineModel> erreursSaisiesRemplissageTempsCollabParSemaine = new();
        //    int nbreHeureMinimumParSemaineEtParCollab = _configurationApp.WebTTTInfoConfig.NbreHeureTotaleMinimumAdeclarerParCollabEtParSemaine;
        //    erreursSaisiesRemplissageTempsCollabParSemaine.AddRange(from KeyValuePair<string, Dictionary<int, float>> saisieParCollabEtParSemaine in listSaisiesTotale
        //                                                            let heureDeclare = 10f
        //                                                            where heureDeclare < nbreHeureMinimumParSemaineEtParCollab
        //                                                            let erreurRemplissage = new SaisieRemplissageTempsCollabParSemaineModel()
        //                                                            {
        //                                                                Qui = "aa",
        //                                                                TotalHeureDeclaree = heureDeclare,
        //                                                                NumeroDeSemaine = 10
        //                                                            }
        //                                                            select erreurRemplissage);
        //    //foreach (KeyValuePair<string, Dictionary<int, float>> saisieParCollabEtParSemaine in listSaisiesTotale)
        //    //{
        //    //    float heureDeclare = 10f;

        //    //    if (heureDeclare < nbreHeureMinimumParSemaineEtParCollab)
        //    //    {
        //    //        var erreurRemplissage = new ErreurSaisieRemplissageTempsParSemaineModel()
        //    //        {
        //    //            Qui = "aa",
        //    //            HeureDeclaree = heureDeclare,
        //    //            NumeroDeSemaine = 10
        //    //        };
        //    //        erreursSaisiesRemplissageTempsCollabParSemaine.Add(erreurRemplissage);
        //    //    }
        //    //}
        //    return erreursSaisiesRemplissageTempsCollabParSemaine;
        //}

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

        private bool IsSaisieValid(ImportWebTTTExcelModel saisieInWebTTT)
        {
            bool result = true;

            if (saisieInWebTTT.Activite == AppliConstant.LblActiviteSpecification)
            {
                _configurationApp.WebTTTInfoConfig.ReglesSaisiesAutorisesParActivite["Spec"].ForEach(mask =>
                {
                    if (saisieInWebTTT.Activite.Contains(mask.Rule))
                    {
                        result = false;
                    }
                });
            }
            else if (saisieInWebTTT.Activite.Equals(AppliConstant.LblActiviteDev) || saisieInWebTTT.Activite.Equals(AppliConstant.LblActiviteQual))
            {
                _configurationApp.WebTTTInfoConfig.ReglesSaisiesAutorisesParActivite["DevQual"].ForEach(mask =>
                {
                    if (saisieInWebTTT.Activite.Contains(mask.Rule))
                    {
                        result = false;
                    }
                });

                if (result)
                {
                    result = (saisieInWebTTT.Application == AppliConstant.LblApplicationParDefaut);
                }
            }

            return result;
        }

        //private void GenereExportCSVErreurSaisies(Dictionary<string, ErreurSaisieDemandeModel> erreursSaisiesDemandes)
        //{
        //    List<CleValeur<string, object>> dataList = new ();
        //    foreach (var demande in erreursSaisiesDemandes)
        //    {
        //        dataList.Add(new CleValeur<string, object>(demande.Key, demande.Value));
        //    }
        //    CSVHelper.GenerateCSVFile(_configurationApp.WebTTTInfoConfig.FileBilanErreurCSV, dataList, false);

        //}

        public void GenereExportCSVErreurSaisies(IList<ErreurSaisieDemandeModel> erreursSaisiesDemandes)
        {
            CSVHelper.GenerateCSVFile(_configurationApp.WebTTTInfoConfig.FileBilanErreurCSV, erreursSaisiesDemandes, false);
        }

        public void GenereExportCSVErreurTempsConsomme(IList<SaisieRemplissageTempsCollabParSemaineModel> tempsConsommes)
        {
            var erreurDeTEmps = tempsConsommes.Where(tpsCollabParSemaine =>
                                                                            tpsCollabParSemaine.TotalHeureDeclaree < _configurationApp.WebTTTInfoConfig.NbreHeureTotaleMinimumAdeclarerParCollabEtParSemaine);


            CSVHelper.GenerateCSVFile(_configurationApp.WebTTTInfoConfig.FileBilanErreurCSV, tempsConsommes, false);
        }

       
    }
}