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
            List<ErreurSaisieFormatDemandeModel> erreursSaisiesDemandes = new();

            List<SaisieRemplissageTempsCollabParSemaineModel> saisiesRemplissageTempsCollabParSemaine = new();

            foreach (ImportWebTTTExcelModel dataRowWebTTT in allDataInWebTTT)
            {
                bool isDemandeValid = InfoSprint.IsDemandeValid(dataRowWebTTT, _configurationApp.WebTTTInfoConfig.ReglesSaisiesAutorisesParActivite);

                if (!isDemandeValid)
                {
                    AddListErreursFormatDeSaisiesDemandes(erreursSaisiesDemandes, dataRowWebTTT);
                }
                AddListTempsDeclareParCollabEtParSemaine(saisiesRemplissageTempsCollabParSemaine, dataRowWebTTT);
            }
            return new ResultImportWebTTT()
            {
                ErreursSaisiesDemandes = erreursSaisiesDemandes,
                SaisiesRemplissageTempsCollabParSemaine = saisiesRemplissageTempsCollabParSemaine,
            };
        }

        //TODO : A faire dans une autre issue
        public Dictionary<string, TempsConsommeDemandeModel> GetTempsConsommeesParDemandeSurUneSemaine(IList<ImportWebTTTExcelModel> saisiesActivitesInWebTTT)
        {
            var result = new Dictionary<string, TempsConsommeDemandeModel>();
            var saisiesActivitesSurLeSprint = saisiesActivitesInWebTTT.Where(data => data.NumeroDeSemaineDateActivite.Equals(_configurationApp.SuiviSprintInfoConfig.NumeroSemaineDebutDeSprint) ||
             data.NumeroDeSemaineDateActivite.Equals(_configurationApp.SuiviSprintInfoConfig.NumeroSemaineFinDeSprint));

            foreach (var saisieActivite in saisiesActivitesSurLeSprint)
            {
                string numeroFormatSuiviSprint = InfoSprint.ModifyDemandeWebTTTToSuiviSprint(saisieActivite.NumeroDeDemande, saisieActivite.Activite);
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
                            IsDemandeValide = InfoSprint.IsDemandeValid(saisieActivite, _configurationApp.WebTTTInfoConfig.ReglesSaisiesAutorisesParActivite)
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

        private static void AddListErreursFormatDeSaisiesDemandes(IList<ErreurSaisieFormatDemandeModel> erreursSaisiesDemandes, ImportWebTTTExcelModel dataWebTTT)
        {
            if (!erreursSaisiesDemandes.Any(err => err.Equals(dataWebTTT.NumeroDeDemande)))
            {
                ErreurSaisieFormatDemandeModel erreurSaisieDemande = new()
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

        public void GenereExportCSVErreurSaisies(IList<ErreurSaisieFormatDemandeModel> erreursSaisiesDemandes)
        {
            if (_configurationApp.WebTTTInfoConfig.TopLaunchBilans)
            {
                if (null !=erreursSaisiesDemandes)
                {
                    CSVHelper.GenerateCSVFile(_configurationApp.WebTTTInfoConfig.FileBilansErreurFormatSaisieDemande, erreursSaisiesDemandes, false);
                }
                else
                {
                    Divers.DisplayInfoMessageInConsole("Aucune erreur de temps de saisie pour les semaines a été détecté");
                }
                
            }
            
        }

        public void GenereBilanErreurTempsConsommeSemaine(IList<SaisieRemplissageTempsCollabParSemaineModel> tempsConsommesParCollab)
        {
            if (_configurationApp.WebTTTInfoConfig.TopLaunchBilans)
            {
                IList<SemaineAvecJourFerieModel> jourFerie = GetListJoursFeries();
                foreach (var tempsConso in tempsConsommesParCollab)
                {
                    int nbreJourSemaine = _configurationApp.WebTTTInfoConfig.NbreJourSemaine;
                    int nbrejourFerieDansLaSemaine = jourFerie.Where(j => j.NumeroDeSemaine.Equals(tempsConso.NumeroDeSemaine)).Select(x => x.NombreDeJourFerie)?.FirstOrDefault() ?? 0;

                    tempsConso.TotalHeureTempsPlein = _configurationApp.WebTTTInfoConfig.NbreHeureTotaleActiviteJour * (nbreJourSemaine - nbrejourFerieDansLaSemaine);
                }

                var erreurDeTEmps = tempsConsommesParCollab.Where(tpsCollabParSemaine =>
                                                                                tpsCollabParSemaine.TotalHeureDeclaree < tpsCollabParSemaine.TotalHeureTempsPlein).ToList();

                if (null != erreurDeTEmps)
                {
                    CSVHelper.GenerateCSVFile(_configurationApp.WebTTTInfoConfig.FileBilansErreurTempsSaisieSemaine, erreurDeTEmps, false);
                }
                else
                {
                    Divers.DisplayInfoMessageInConsole("Aucune erreur de temps de saisie pour les semaines a été détecté");
                }
            }
            
        }

        private IList<SemaineAvecJourFerieModel> GetListJoursFeries()
        {
            var listJoursFeries = _configurationApp.WebTTTInfoConfig.JoursFeries;
            var listJourFerieAvecNumSemaine = new List<SemaineAvecJourFerieModel>();
            foreach (var jour in listJoursFeries)
            {
                var numSemaineDuJourFerie = InfoSprint.GetNumSemaine(DateTime.Parse(jour.Value));
                if (listJourFerieAvecNumSemaine.Any(j => j.NumeroDeSemaine.Equals(numSemaineDuJourFerie)))
                {
                    var semaineFerie = listJourFerieAvecNumSemaine.FirstOrDefault(j => j.NumeroDeSemaine.Equals(numSemaineDuJourFerie));
                    if (null !=semaineFerie)
                    {
                        semaineFerie.NombreDeJourFerie += 1;
                    }
                }
                else
                {
                    var jourFerie = new SemaineAvecJourFerieModel()
                    {
                        NombreDeJourFerie = 1,
                        NumeroDeSemaine = numSemaineDuJourFerie
                    };
                    listJourFerieAvecNumSemaine.Add(jourFerie);
                }
            }
            return listJourFerieAvecNumSemaine;
        }
    }
}