using MajConsoSuiviSprint.Cli.Business.Interfaces;
using MajConsoSuiviSprint.Cli.Constants;
using MajConsoSuiviSprint.Cli.Model;
using MajConsoSuiviSprint.Cli.Utils;
using Microsoft.Extensions.Configuration;

namespace MajConsoSuiviSprint.Cli.Business
{
    internal class ConfigurationsApp : IConfigurationsApp
    {
      
        public WebTTTInfoConfigModel WebTTTInfoConfig { get; set; } = default!;
        public SuiviSprintInfoConfigModel SuiviSprintInfoConfig { get; set; } = default!;

        public ConfigurationsApp(string pathAppSettings)
        {

            Console.WriteLine("Configuration");


            string jsonFile = Divers.GetFileNameFromFullPathFilename(pathAppSettings);
            string pathJson = Divers.GetPathFromFullPathFilename(pathAppSettings);

            IConfiguration config = new ConfigurationBuilder()
                                    .SetBasePath(pathJson)
                                    .AddJsonFile(jsonFile, optional: false, reloadOnChange: true)
                                    .Build();
           

            WebTTTInfoConfig = LoadInfosWebTTTFromSettings(config);
            SuiviSprintInfoConfig = LoadInfosSuiviSprintFromSettings(config);
            InitInfosSuiviSprint();
            InitWebTTT();

            CheckCoherenceParamNumDeSemaineAPartirDe();
        }

        private void CheckCoherenceParamNumDeSemaineAPartirDe()
        {
            if (WebTTTInfoConfig.NumeroDeSemaineAPartirDuquelChecker > SuiviSprintInfoConfig.NumeroSemaineDebutDeSprint)
            {
                throw new Exception($"Erreur de paramétrage. Le paramètre NumeroDeSemaineAPartirDuquelChecker ne peut être supérieur à la semaine du fichier de suivi {SuiviSprintInfoConfig.NumeroSemaineDebutDeSprint}");
            }
        }

        private static WebTTTInfoConfigModel LoadInfosWebTTTFromSettings(IConfiguration config)
        {

            var sectionWebTTT = config.GetSection(AppliConstant.WebTTTSection);
            var webTTTInfo = new WebTTTInfoConfigModel()
            {
                FileBilanErreurCSV = sectionWebTTT.GetValue<string>("FileNameBilanErreurSaisieDansWebTTT") ?? default!,
                Path = sectionWebTTT.GetValue<string>("Path") ?? default!,
                FileName = sectionWebTTT.GetValue<string>("FileName") ?? default!,
                SheetName = sectionWebTTT.GetValue<string>("SheetName") ?? default!,
                NbreDeSemaineAPrendreAvtLaSemaineEnCours = sectionWebTTT.GetValue<int>("NbreDeSemaineAPrendreAvtLaSemaineEnCours"),
                NbreHeureTotaleMinimumAdeclarerParCollabEtParSemaine = sectionWebTTT.GetValue<int>("NbreHeureTotaleMinimumAdeclarerParCollabEtParSemaine"),
                Headers = sectionWebTTT
                                .GetSection("Headers")
                                .Get<List<HeadersWebTTTModel>>()
                                ?.AsReadOnly()
                                ?? new List<HeadersWebTTTModel>().AsReadOnly()
            };
            List<MaskSaisieModel> maskSpecAutorise = sectionWebTTT
                                .GetSection("MaskSaisieModelDemande").GetSection("Spec")
                                .Get<List<MaskSaisieModel>>()
                                ?? new List<MaskSaisieModel>()
                               ;

            webTTTInfo.ReglesSaisiesAutorisesParActivite.Add("Spec", maskSpecAutorise);

            List<MaskSaisieModel> maskdevQualAutorise = sectionWebTTT
                                .GetSection("MaskSaisieModelDemande").GetSection("DevQual")
                                .Get<List<MaskSaisieModel>>()
                                ?? new List<MaskSaisieModel>();

            webTTTInfo.ReglesSaisiesAutorisesParActivite.Add("DevQual", maskdevQualAutorise);
            return webTTTInfo;
        }


        private static SuiviSprintInfoConfigModel LoadInfosSuiviSprintFromSettings(IConfiguration config)
        {
            var sectionSuiviSprint = config.GetSection(AppliConstant.SuiviSprintSection);

            var suiviSprintInfo = new SuiviSprintInfoConfigModel()
            {
                FileName = sectionSuiviSprint.GetValue<string>("FileName") ?? default!,
                Path = sectionSuiviSprint.GetValue<string>("Path") ?? default!,
                MajSuiviSprint = sectionSuiviSprint.GetValue<string>("MajSuiviSprint")?.ToUpper() ?? default!,
            };

            suiviSprintInfo.TabSuivi.SheetName = sectionSuiviSprint
                                                .GetSection("TabSuivi")
                                                .GetValue<string>("SheetName") ?? default!;

            suiviSprintInfo.TabSuivi.TableName = sectionSuiviSprint
                                                .GetSection("TabSuivi")
                                                .GetValue<string>("TableName") ?? default!;
            int? numColonne = int.Parse(sectionSuiviSprint.GetSection("TabSuivi")
                                    .GetSection("NumColumnTable")
                                    .GetValue<string>("NoColumnApplication") ?? default!);
            suiviSprintInfo.TabSuivi.NumColumnTable.NoColumnApplication = numColonne ?? default;

            numColonne = int.Parse(sectionSuiviSprint
                            .GetSection("TabSuivi")
                            .GetSection("NumColumnTable")
                            .GetValue<string>("NoColumnDemande") ?? default!);
            suiviSprintInfo.TabSuivi.NumColumnTable.NoColumnDemande = numColonne ?? default;

            numColonne = int.Parse(sectionSuiviSprint
                            .GetSection("TabSuivi")
                            .GetSection("NumColumnTable")
                            .GetValue<string>("NoColumnHoursDevConsumed") ?? default!);
            suiviSprintInfo.TabSuivi.NumColumnTable.NoColumnHoursDevConsumed = numColonne ??default;
            return suiviSprintInfo;
        }

        private void InitInfosSuiviSprint()
        {

            string anneeEC = DateTime.Now.Year.ToString();
            if (string.IsNullOrEmpty(SuiviSprintInfoConfig.FileName))
            {
                SuiviSprintInfoConfig.FileName = InfoSprint.GetFileNameSuiviSprintEC();
            }
            else
            {
                SuiviSprintInfoConfig.FileName = SuiviSprintInfoConfig.FileName.Replace("{anneeEC}", anneeEC);
            }

            SuiviSprintInfoConfig.Path = SuiviSprintInfoConfig.Path
                                       .Replace("{anneeEC}", anneeEC)
                                       .Replace("{userName}", Environment.UserName);

            bool isCheminAvecBackSlashALaFin = IsPathWithBackSlash(SuiviSprintInfoConfig.Path);

            SuiviSprintInfoConfig.FullFileName = $@"{SuiviSprintInfoConfig.Path}{(!isCheminAvecBackSlashALaFin ? "\\" : "")}{SuiviSprintInfoConfig.FileName}";

        }

        private void InitWebTTT()
        {
            Console.WriteLine("InitWebTTT");
            string anneeEC = DateTime.Now.Year.ToString();
            WebTTTInfoConfig.FileName = WebTTTInfoConfig.FileName.Replace("{anneeEC}", anneeEC);
            WebTTTInfoConfig.Path = WebTTTInfoConfig.Path
                                        .Replace("{anneeEC}", anneeEC)
                                        .Replace("{userName}", Environment.UserName);



            WebTTTInfoConfig.SheetName = WebTTTInfoConfig.SheetName
                                            .Replace("{anneeEC}", DateTime.Now.Year.ToString());

            if (Directory.Exists(WebTTTInfoConfig.Path))
            {
                bool isCheminAvecBackSlashALaFin = IsPathWithBackSlash(WebTTTInfoConfig.Path);

                WebTTTInfoConfig.FullFileName = $@"{WebTTTInfoConfig.Path}{(!isCheminAvecBackSlashALaFin ? "\\" : "")}{WebTTTInfoConfig.FileName}";
            }
        }

        private static bool IsPathWithBackSlash(string path)
        {
            return path.EndsWith(@"\");
        }
    }
}
