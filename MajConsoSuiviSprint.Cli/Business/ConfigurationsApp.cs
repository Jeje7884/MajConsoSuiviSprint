using MajConsoSuiviSprint.Cli.Business.Interfaces;
using MajConsoSuiviSprint.Cli.Model;
using MajConsoSuiviSprint.Cli.Utils;
using Microsoft.Extensions.Configuration;

namespace MajConsoSuiviSprint.Cli.Business
{
    internal class ConfigurationsApp : IConfigurationsApp
    {
        private const string WebTTTSection = "WebTTT";
        private const string SuiviSprintSection = "SuiviSprint";
        private string JsonFile = "appSettings.json";
        private string PathJson= Directory.GetCurrentDirectory();

        public WebTTTInfoModel WebTTTModel { get; set; } = default!;
        public SuiviSprintModel SuiviSprintModel { get; set; } = default!;

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

            WebTTTInfoModel webTTTInfo = LoadInfosWebTTTFromSettings(config);
            SuiviSprintModel = LoadInfosSuiviSprintFromSettings(config);

            InitWebTTT(ref webTTTInfo);
            WebTTTModel = webTTTInfo;

        }

        private static WebTTTInfoModel LoadInfosWebTTTFromSettings(IConfiguration config)
        {

            var sectionWebTTT = config.GetSection(WebTTTSection);
            var webTTTInfo = new WebTTTInfoModel()
            {
                FileBilanErreurCSV = sectionWebTTT.GetValue<string>("FileNameBilanErreurSaisieDansWebTTT") ?? "",
                Path = sectionWebTTT.GetValue<string>("Path") ?? "",
                FileName = sectionWebTTT.GetValue<string>("FileName") ?? "",
                SheetName = sectionWebTTT.GetValue<string>("SheetName") ?? "",
                NbreSprintAPrendreEnCompte = sectionWebTTT.GetValue<int>("NbreSprintAPrendreEnCompte"),
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


        private static SuiviSprintModel LoadInfosSuiviSprintFromSettings(IConfiguration config)
        {
            var sectionSuiviSprint = config.GetSection(SuiviSprintSection);

            var suiviSprintInfo = new SuiviSprintModel()
            {
                FileName = sectionSuiviSprint.GetValue<string>("FolderName") ?? "",
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

        private static void InitWebTTT(ref WebTTTInfoModel WebTTTFile)
        {
            Console.WriteLine("InitWebTTT");
            string anneeEC = DateTime.Now.Year.ToString();
            string username = Environment.UserName;

            if (string.IsNullOrEmpty(WebTTTFile.FileName))
            {
                WebTTTFile.FileName = InfoSprint.GetFileNameSuiviSprintEC();
            }
            else
            {
                WebTTTFile.FileName = WebTTTFile.FileName.Replace("{anneeEC}", anneeEC);
            }
            WebTTTFile.Path = WebTTTFile.Path
                                        .Replace("{anneeEC}", anneeEC)
                                        .Replace("{userName}", username);



            WebTTTFile.SheetName = WebTTTFile.SheetName.Replace("{anneeEC}", DateTime.Now.Year.ToString());

            if (Directory.Exists(WebTTTFile.Path))
            {
                WebTTTFile.FullFileName = $@"{WebTTTFile.Path}\{WebTTTFile.FileName}";
            }
        }
    }
}
