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

        internal List<ImportWebTTTExcelModel> ImportInfosFromWebTTT()
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

        public ResultImportWebTTT CheckSaisiesActiviteInWebTTT(List<ImportWebTTTExcelModel> allDataInWebTTT)
        {
            //Dictionary<string, ErreurSaisieDemandeModel> erreursSaisiesDemandes = new();
            List<ErreurSaisieDemandeModel> erreursSaisiesDemandes = new();

            bool alreadyExists = erreursSaisiesDemandes.Any(x => x.NumeroDeDemande == "ooo");
            if (erreursSaisiesDemandes.Exists(x => x.NumeroDeDemande == "200"))
            {
                //code
            }

            //Dictionary<string, TempsConsommeDemandeModel> tempsConsommesPardemandesEtParSprint = new();

            //Dictionary<string, Dictionary<int, float>> tempsDeclaresParCollab = new();
            List<SaisieRemplissageTempsCollabParSemaineModel> erreursSaisiesRemplissageTempsCollabParSemaine = new();

            foreach (ImportWebTTTExcelModel dataWebTTT in allDataInWebTTT)
            {
                bool isDemandeValid = IsSaisieValid(dataWebTTT);

                if (!isDemandeValid)
                {
                    AddListErreursDeSaisies(erreursSaisiesDemandes, dataWebTTT);
                }
                //if (SaisieAPrendreEnComptePourTempsConsoFichierDeSuivi(dataWebTTT))
                //{
                //    AddDictionnaireTempsConsommeParDemandeEtParPole(tempsConsommesPardemandesEtParSprint, dataWebTTT);

                //    tempsConsommesPardemandesEtParSprint[dataWebTTT.NumeroDeDemande].IsDemandeValide = isDemandeValid;
                //}
                AddDictionnaireTempsDeclareParCollabEtParSemaine(erreursSaisiesRemplissageTempsCollabParSemaine, dataWebTTT);
                //erreursSaisiesRemplissageTempsCollabParSemaine = GetListCollabErreurRemplissageHeruresWebTTTT(tempsDeclaresParCollab);
            }
            return new ResultImportWebTTT()
            {
                ErreursSaisiesDemandes = erreursSaisiesDemandes,
                ErreursSaisiesRemplissageTempsCollabParSemaine = erreursSaisiesRemplissageTempsCollabParSemaine,
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

        private static void AddDictionnaireTempsDeclareParCollabEtParSemaine(List<SaisieRemplissageTempsCollabParSemaineModel> tempsDeclaresParCollab, ImportWebTTTExcelModel dataWebTTT)
        {
            //if (!tempsDeclaresParCollab.ContainsKey(dataWebTTT.TrigrammeCollab))
            //{
            //    var tempsdeclareCollab = new Dictionary<int, float>
            //        {
            //            { dataWebTTT.NumeroDeSemaineDateActivite, dataWebTTT.HeureDeclaree }
            //        };
            //    tempsDeclaresParCollab.Add(dataWebTTT.TrigrammeCollab, tempsdeclareCollab);
            //}
            //else
            //{
            //    if (tempsDeclaresParCollab[dataWebTTT.TrigrammeCollab].ContainsKey(dataWebTTT.NumeroDeSemaineDateActivite))
            //    {
            //        tempsDeclaresParCollab[dataWebTTT.TrigrammeCollab][dataWebTTT.NumeroDeSemaineDateActivite] += dataWebTTT.HeureDeclaree;
            //    }
            //    else
            //    {
            //        tempsDeclaresParCollab[dataWebTTT.TrigrammeCollab].Add(dataWebTTT.NumeroDeSemaineDateActivite, dataWebTTT.HeureDeclaree);
            //    }
            //}
        }



        private static void AddListErreursDeSaisies(List<ErreurSaisieDemandeModel> erreursSaisiesDemandes, ImportWebTTTExcelModel dataWebTTT)
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

        public void GenereExportCSVErreurSaisies(List<ErreurSaisieDemandeModel> erreursSaisiesDemandes)
        {
            CSVHelper.GenerateCSVFile(_configurationApp.WebTTTInfoConfig.FileBilanErreurCSV, erreursSaisiesDemandes, false);
        }

        private bool SaisieAPrendreEnComptePourTempsConsoFichierDeSuivi(ImportWebTTTExcelModel saisieInWebTTT)
        {
            return (InfoSprint.IsActivityToManaged(saisieInWebTTT.Activite) &&
                    InfoSprint.IsPeriodeToManaged(saisieInWebTTT.DateDeSaisie,
                                                _configurationApp.SuiviSprintInfoConfig.NumeroSemaineDebutDeSprint,
                                                _configurationApp.SuiviSprintInfoConfig.NumeroSemaineFinDeSprint));
        }
    }
}