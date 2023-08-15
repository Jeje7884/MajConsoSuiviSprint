using MajConsoSuiviSprint.Cli.Business.Interfaces;
using MajConsoSuiviSprint.Cli.Model;
using MajConsoSuiviSprint.Cli.Utils;
using Microsoft.Extensions.Configuration;
using Org.BouncyCastle.Bcpg;

namespace MajConsoSuiviSprint.Cli.Business
{
    internal class ConfigurationsApp : IConfigurationsApp
    {
        private const string WebTTTSection = "WebTTT";
        private const string SuiviSprintSection = "SuiviSprint";
        private readonly string JsonFile = "appSettings.json";
        private readonly string PathJson = Directory.GetCurrentDirectory();

        public WebTTTInfoConfigModel WebTTTInfoConfig { get; set; } = default!;
        public SuiviSprintInfoConfigModel SuiviSprintInfoConfig { get; set; } = default!;

        public ConfigurationsApp(string pathAppSettings)
        {
           
            Console.WriteLine("Configuration");

            if (!string.IsNullOrEmpty(pathAppSettings))
            {
                JsonFile= Divers.GetFileNameFromFullPathFilename(pathAppSettings);
                PathJson= Divers.GetPathFromFullPathFilename(pathAppSettings);
            }
            IConfiguration config = new ConfigurationBuilder()
                                    .SetBasePath(PathJson)
                                    .AddJsonFile(JsonFile, optional: false, reloadOnChange: true)
                                    .Build();

            WebTTTInfoConfig = LoadInfosWebTTTFromSettings(config);
            SuiviSprintInfoConfig = LoadInfosSuiviSprintFromSettings(config);
            InitInfosSuiviSprint();
            InitWebTTT();
            //WebTTTInfoConfig = webTTTInfo;

        }

        private static WebTTTInfoConfigModel LoadInfosWebTTTFromSettings(IConfiguration config)
        {

            var sectionWebTTT = config.GetSection(WebTTTSection);
            var webTTTInfo = new WebTTTInfoConfigModel()
            {
                FileBilanErreurCSV = sectionWebTTT.GetValue<string>("FileNameBilanErreurSaisieDansWebTTT") ?? "",
                Path = sectionWebTTT.GetValue<string>("Path") ?? "",
                FileName = sectionWebTTT.GetValue<string>("FileName") ?? "",
                SheetName = sectionWebTTT.GetValue<string>("SheetName") ?? "",
                NbreSprintAPrendreEnCompte = sectionWebTTT.GetValue<int>("NbreSprintAPrendreEnCompte"),
                NbreHeureTotaleMinimumAdeclarerParCollabEtParSemaine = sectionWebTTT.GetValue<int>("NbreHeureTotaleMinimumAdeclarerParCollabEtParSemaine"),
                Headers = sectionWebTTT
                                .GetSection("Headers")
                                .Get<List<HeadersWebTTTModel>>()
                                ?.AsReadOnly()
                                ?? new List<HeadersWebTTTModel>().AsReadOnly()
            };
            var maskSpecAutorise = sectionWebTTT
                                .GetSection("MaskSaisieModelDemande").GetSection("Spec")
                                .Get<List<MaskSaisieModel>>()
                                ?? new List<MaskSaisieModel>()   
                               ;

            webTTTInfo.ReglesSaisiesAutorisesParActivite.Add("Spec", maskSpecAutorise);

            var maskdevQualAutorise = sectionWebTTT
                                .GetSection("MaskSaisieModelDemande").GetSection("DevQual")
                                .Get<List<MaskSaisieModel>>()
                                ?? new List<MaskSaisieModel>();

            webTTTInfo.ReglesSaisiesAutorisesParActivite.Add("DevQual", maskdevQualAutorise);
            return webTTTInfo;
        }


        private static SuiviSprintInfoConfigModel LoadInfosSuiviSprintFromSettings(IConfiguration config)
        {
            var sectionSuiviSprint = config.GetSection(SuiviSprintSection);

            var suiviSprintInfo = new SuiviSprintInfoConfigModel()
            {
                FileName = sectionSuiviSprint.GetValue<string>("FileName") ?? "",
                Path = sectionSuiviSprint.GetValue<string>("Path") ?? ""
            };

            suiviSprintInfo.TabSuivi.SheetName = sectionSuiviSprint
                                                .GetSection("TabSuivi")
                                                .GetValue<string>("SheetName") ?? "";

            suiviSprintInfo.TabSuivi.TableName = sectionSuiviSprint
                                                .GetSection("TabSuivi")
                                                .GetValue<string>("TableName") ?? "";
            int? numColonne = int.Parse(sectionSuiviSprint.GetSection("TabSuivi")
                                    .GetSection("NumColumnTable")
                                    .GetValue<string>("NoColumnApplication") ?? "");
            suiviSprintInfo.TabSuivi.NumColumnTable.NoColumnApplication = numColonne ?? 0;

            numColonne = int.Parse(sectionSuiviSprint
                            .GetSection("TabSuivi")
                            .GetSection("NumColumnTable")
                            .GetValue<string>("NoColumnDemande") ?? "");
            suiviSprintInfo.TabSuivi.NumColumnTable.NoColumnDemande = numColonne ?? 0;

            numColonne = int.Parse(sectionSuiviSprint
                            .GetSection("TabSuivi")
                            .GetSection("NumColumnTable")
                            .GetValue<string>("NoColumnHoursDevConsumed") ?? "");
            suiviSprintInfo.TabSuivi.NumColumnTable.NoColumnHoursDevConsumed = numColonne ?? 0;
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



            WebTTTInfoConfig.SheetName = WebTTTInfoConfig.SheetName.Replace("{anneeEC}", DateTime.Now.Year.ToString());

            if (Directory.Exists(WebTTTInfoConfig.Path))
            {
                bool isCheminAvecBackSlashALaFin = IsPathWithBackSlash(WebTTTInfoConfig.Path);

                //string fullPathWebbTTTFile = _configurationApp.WebTTTInfoConfig.Path + (!isCheminAvecBackSlashALaFin ? "\\" : "") + _configurationApp.WebTTTInfoConfig.FileName;
                WebTTTInfoConfig.FullFileName = $@"{WebTTTInfoConfig.Path}{(!isCheminAvecBackSlashALaFin ? "\\" : "")}{WebTTTInfoConfig.FileName}";
            }

            (int debutSprint, int finSprint) semainesSprint = InfoSprint.ExtractSemainesSprint(SuiviSprintInfoConfig.FileName);

            WebTTTInfoConfig.NumeroDebutSemaineAImporter = semainesSprint.debutSprint;
            WebTTTInfoConfig.NumeroFinSemaineAImporter = semainesSprint.finSprint;


        }

        private bool IsPathWithBackSlash(string path)
        {
            return path.EndsWith("\\");
        }
    }
}
