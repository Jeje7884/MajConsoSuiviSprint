using MajConsoSuiviSprint.Cli.Business.Interfaces;
using MajConsoSuiviSprint.Cli.Constants;
using MajConsoSuiviSprint.Cli.Helper;
using MajConsoSuiviSprint.Cli.Model;
using MajConsoSuiviSprint.Cli.Utils;

namespace MajConsoSuiviSprint.Cli.Business
{
    internal class ImportWebTTT
    {
        //public List<ImportWebTTTExcelModel> ImportSaisisWebTTTExcel = new();
        //public List<ResultatImport> ResultatsWebTTTExcel = new();
        //public List<ErreurSaisieDemandeModel> ErreursDeSaisies = new();
        private readonly IConfigurationsApp _configurationApp;

        public ImportWebTTT(IConfigurationsApp configuration)
        {
            _configurationApp = configuration;
        }

        //internal class TimeConsumedByActivite
        //{
        //    public int HourQual { get; set; } = 0;

        //    public int HourDev { get; set; } = 0;
        //    public string Application { get; set; } = string.Empty;
        //    public bool NumDemandeIsValid { get; set; } = true;
        //}

        //internal class ResultatImport
        //{
        //    public List<ErreurSaisieDemandeModel> ErrorsSaisis { get; set; } = new();
        //    public Dictionary<string, TimeConsumedByActivite> TempsConsommesParDemandes { get; set; } = new();

        //}

        internal ResultImportWebTTT ImportInfosFromWebTTT()
        {
            if (Divers.IsFileOpened(_configurationApp.WebTTTInfoConfig.FullFileName))
            {
                throw new ExceptFileOpenException($"Le fichier {_configurationApp.WebTTTInfoConfig.FullFileName} est ouvert");
            }
            var result = ExceLNPOIHelper.ImportFichierWebTTTExcel(_configurationApp.WebTTTInfoConfig);
            Console.WriteLine($" {result.Count} lignes ont été récupérés du fichier {_configurationApp.WebTTTInfoConfig.FileName}");

            return new ResultImportWebTTT();
        }

        private ResultImportWebTTT CheckSaisiesActiviteInWebTTT(List<ImportWebTTTExcelModel> allDataInWebTTT)
        {
            Dictionary<string, ErreurSaisieDemandeModel> erreursSaisiesDemandes = new();
            Dictionary<string, TempsConsommeDemandeModel> tempsConsommesPardemandesEtParSprint = new();
            Dictionary<string, Dictionary<int, float>> tempsDeclaresParCollab = new();
            List<ErreurSaisieRemplissageTempsParSemaineModel> erreursSaisiesRemplissageTempsCollabParSemaine = new();

            foreach (ImportWebTTTExcelModel dataWebTTT in allDataInWebTTT)
            {
                bool isDemandeValid = IsSaisieValid(dataWebTTT);

                if (!isDemandeValid)
                {
                    AddDictionnaireErreursDeSaisies(erreursSaisiesDemandes, dataWebTTT);
                }
                if (SaisieAPrendreEnComptePourTempsConsoFichierDeSuivi(dataWebTTT))
                {
                    AddDictionnaireTempsConsommeParDemandeEtParPole(tempsConsommesPardemandesEtParSprint, dataWebTTT);

                    tempsConsommesPardemandesEtParSprint[dataWebTTT.NumeroDeDemande].IsDemandeValide = isDemandeValid;
                }
                AddDictionnaireTempsDeclareParCollabEtParSemaine(tempsDeclaresParCollab, dataWebTTT);
                erreursSaisiesRemplissageTempsCollabParSemaine = GetListCollabErreurRemplissageHeruresWebTTTT(tempsDeclaresParCollab);
            }
            return new ResultImportWebTTT()
            {
                ErreursSaisiesDemandes = erreursSaisiesDemandes,
                ErreursSaisiesRemplissageTempsCollabParSemaine = erreursSaisiesRemplissageTempsCollabParSemaine,
                TempsConsommesPardemandesEtParSprint = tempsConsommesPardemandesEtParSprint
            };
        }

        private List<ErreurSaisieRemplissageTempsParSemaineModel> GetListCollabErreurRemplissageHeruresWebTTTT(Dictionary<string, Dictionary<int, float>> listSaisiesTotale)
        {
            List<ErreurSaisieRemplissageTempsParSemaineModel> erreursSaisiesRemplissageTempsCollabParSemaine = new();
            int nbreHeureMinimumParSemaineEtParCollab = _configurationApp.WebTTTInfoConfig.NbreHeureTotaleMinimumAdeclarerParCollabEtParSemaine;
            erreursSaisiesRemplissageTempsCollabParSemaine.AddRange(from KeyValuePair<string, Dictionary<int, float>> saisieParCollabEtParSemaine in listSaisiesTotale
                                                                    let heureDeclare = 10f
                                                                    where heureDeclare < nbreHeureMinimumParSemaineEtParCollab
                                                                    let erreurRemplissage = new ErreurSaisieRemplissageTempsParSemaineModel()
                                                                    {
                                                                        Qui = "aa",
                                                                        HeureDeclaree = heureDeclare,
                                                                        NumeroDeSemaine = 10
                                                                    }
                                                                    select erreurRemplissage);
            //foreach (KeyValuePair<string, Dictionary<int, float>> saisieParCollabEtParSemaine in listSaisiesTotale)
            //{
            //    float heureDeclare = 10f;

            //    if (heureDeclare < nbreHeureMinimumParSemaineEtParCollab)
            //    {
            //        var erreurRemplissage = new ErreurSaisieRemplissageTempsParSemaineModel()
            //        {
            //            Qui = "aa",
            //            HeureDeclaree = heureDeclare,
            //            NumeroDeSemaine = 10
            //        };
            //        erreursSaisiesRemplissageTempsCollabParSemaine.Add(erreurRemplissage);
            //    }
            //}
            return erreursSaisiesRemplissageTempsCollabParSemaine;
        }

        private static void AddDictionnaireTempsDeclareParCollabEtParSemaine(Dictionary<string, Dictionary<int, float>> tempsDeclaresParCollab, ImportWebTTTExcelModel dataWebTTT)
        {
            if (!tempsDeclaresParCollab.ContainsKey(dataWebTTT.TrigrammeCollab))
            {
                var tempsdeclareCollab = new Dictionary<int, float>
                    {
                        { dataWebTTT.NumeroDeSemaineDateActivite, dataWebTTT.HeureDeclaree }
                    };
                tempsDeclaresParCollab.Add(dataWebTTT.TrigrammeCollab, tempsdeclareCollab);
            }
            else
            {
                if (tempsDeclaresParCollab[dataWebTTT.TrigrammeCollab].ContainsKey(dataWebTTT.NumeroDeSemaineDateActivite))
                {
                    tempsDeclaresParCollab[dataWebTTT.TrigrammeCollab][dataWebTTT.NumeroDeSemaineDateActivite] += dataWebTTT.HeureDeclaree;
                }
                else
                {
                    tempsDeclaresParCollab[dataWebTTT.TrigrammeCollab].Add(dataWebTTT.NumeroDeSemaineDateActivite, dataWebTTT.HeureDeclaree);
                }
            }
        }

        private static void AddDictionnaireTempsConsommeParDemandeEtParPole(Dictionary<string, TempsConsommeDemandeModel> tempsConsommesPardemandesEtParSprint, ImportWebTTTExcelModel dataWebTTT)
        {
            if (!tempsConsommesPardemandesEtParSprint.ContainsKey(dataWebTTT.NumeroDeDemande))
            {
                TempsConsommeDemandeModel tempsConsoDemande = new()
                {
                    NumeroDeDemande = dataWebTTT.NumeroDeDemande,
                    Application = dataWebTTT.Application,
                    HeureDeDeveloppement = dataWebTTT.Activite.Equals(AppliConstant.LblActiviteDev) ? dataWebTTT.HeureDeclaree : 0f,
                    HeureDeQualification = dataWebTTT.Activite.Equals(AppliConstant.LblActiviteQual) ? dataWebTTT.HeureDeclaree : 0f
                };
                tempsConsommesPardemandesEtParSprint.Add(dataWebTTT.NumeroDeDemande, tempsConsoDemande);
            }
        }

        private static void AddDictionnaireErreursDeSaisies(Dictionary<string, ErreurSaisieDemandeModel> erreursSaisiesDemandes, ImportWebTTTExcelModel dataWebTTT)
        {
            if (!erreursSaisiesDemandes.ContainsKey(dataWebTTT.NumeroDeDemande))
            {
                ErreurSaisieDemandeModel erreurSaisieDemande = new()
                {
                    Activite = dataWebTTT.Activite,
                    Application = dataWebTTT.Application,
                    NumeroDemande = dataWebTTT.NumeroDeDemande,
                    Qui = dataWebTTT.TrigrammeCollab,
                    DateDeSaisie = dataWebTTT.LblDateDeSaisie,
                    DetailErreur = AppliConstant.MessageErreurSaisie
                };
                erreursSaisiesDemandes.Add(dataWebTTT.NumeroDeDemande, erreurSaisieDemande);
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

        private void GenereExportCSVErreurSaisies(Dictionary<string, ErreurSaisieDemandeModel> erreursSaisiesDemandes)
        {
            //var test= List<CleValeur<int, string>> dataList = erreursSaisiesDemandes
            List<CleValeur<string, object>> dataList = new List<CleValeur<string, object>>();
            foreach (var demande in erreursSaisiesDemandes)
            {
                dataList.Add(new CleValeur<string, object>(demande.Key, demande.Value));
            }
            CSVHelper.GenerateCSVFile(_configurationApp.WebTTTInfoConfig.FileBilanErreurCSV, dataList, false);

            // var test= List<CleValeur< Dictionary<string, ErreurSaisieDemandeModel>>>> dataList = erreursSaisiesDemandes
            //.Select(pair => new CleValeur<int, string>(pair.Key, pair.Value))
            //.ToList();
        }

        private bool SaisieAPrendreEnComptePourTempsConsoFichierDeSuivi(ImportWebTTTExcelModel saisieInWebTTT)
        {
            return (InfoSprint.IsActivityToManaged(saisieInWebTTT.Activite) && InfoSprint.IsPeriodeToManaged(saisieInWebTTT.DateDeSaisie, _configurationApp.SuiviSprintInfoConfig.NumeroSemaineDebutDeSprint, _configurationApp.SuiviSprintInfoConfig.NumeroSemaineFinDeSprint));
        }
    }
}