using MajConsoSuiviSprint.Cli.Model;
using MajConsoSuiviSprint.Cli.Utils;
using Microsoft.Extensions.Configuration;

namespace MajConsoSuiviSprint.Cli.Business
{
    internal class Configuration
    {
        private const string WebTTTSection = "WebTTTFile";
        private const string SuiviSprintSection = "SuiviSprint";

        public string FileBilanErreurCSV { get; set; }
        public string PathSharepointSuiviSprint { get; set; }
        public string PathSharepointSuiviSprint2 { get; set; }
        public WebTTTInfoModel WebTTTModel { get; set; } = default!;
        public SuiviSprintModel SuiviSprintModel { get; set; } = default!;

        public Configuration()
        {
            Console.WriteLine("Configuration");
            PathSharepointSuiviSprint = PathSharepointSuiviSprint2 = FileBilanErreurCSV = string.Empty;

            IConfiguration config = new ConfigurationBuilder()
                                    .SetBasePath(Directory.GetCurrentDirectory())
                                    .AddJsonFile("appSettings.json", optional: false, reloadOnChange: true)
                                    .Build();

            //FileBilanErreurCSV = config.GetSection(nameof(Applications)).Get<List<Application>>().AsReadOnly();
            WebTTTInfoModel webTTTInfo = LoadInfosWebTTTFromSettings(config);
            SuiviSprintModel = LoadInfosSuiviSprintFromSettings(config);

            InitWebTTT(ref webTTTInfo, PathSharepointSuiviSprint, PathSharepointSuiviSprint);
            WebTTTModel = webTTTInfo;

        }

        private static SuiviSprintModel LoadInfosSuiviSprintFromSettings(IConfiguration config)
        {
            var sectionSuiviSprint = config.GetSection(SuiviSprintSection);

            var suiviSprintInfo = new SuiviSprintModel()
            {
                FolderName = sectionSuiviSprint.GetValue<string>("FolderName") ?? "",
                FileName = sectionSuiviSprint.GetValue<string>("FolderName") ?? "",

            };
            suiviSprintInfo.TabSuivi.SheetName = sectionSuiviSprint.GetSection("TabSuivi").GetValue<string>("SheetName") ?? "";
            suiviSprintInfo.TabSuivi.TableName = sectionSuiviSprint.GetSection("TabSuivi").GetValue<string>("TableName") ?? "";
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

        private WebTTTInfoModel LoadInfosWebTTTFromSettings(IConfiguration config)
        {
            FileBilanErreurCSV = config.GetValue<string>("fileNameBilanErreurSaisieDansWebTTT") ?? "";
            PathSharepointSuiviSprint = config.GetValue<string>(nameof(PathSharepointSuiviSprint)) ?? "";
            PathSharepointSuiviSprint2 = config.GetValue<string>(nameof(PathSharepointSuiviSprint2)) ?? "";

            var sectionWebTTT = config.GetSection(WebTTTSection);
            var webTTTInfo = new WebTTTInfoModel()
            {
                FolderName = sectionWebTTT.GetValue<string>("FolderName") ?? "",
                FileName = sectionWebTTT.GetValue<string>("FileName") ?? "",
                SheetName = sectionWebTTT.GetValue<string>("SheetName") ?? "",
                NbreSprintAPrendreEnCompte = sectionWebTTT.GetValue<int>("NbreSprintAPrendreEnCompte"),
                Headers = sectionWebTTT
                                .GetSection("Headers")
                                .Get<List<HeadersWebTTT>>()
                                ?.AsReadOnly()
                                ?? new List<HeadersWebTTT>().AsReadOnly()
            };
            return webTTTInfo;
        }

        private static string InitFileNameBilanErreurSaisieDansWebTTT()
        {
            return "";
        }

        private static void InitWebTTT(ref WebTTTInfoModel WebTTTFile, string pathSharepointSuiviSprint, string pathSharepointSuiviSprint2)
        {
            Console.WriteLine("InitWebTTT");
            if (string.IsNullOrEmpty(WebTTTFile.FileName))
            {
                WebTTTFile.FileName = InfoSprint.GetFileNameSuiviSprintEC();
            }
            else
            {
                WebTTTFile.FileName = WebTTTFile.FileName.Replace("{anneeEC}", DateTime.Now.Year.ToString());
            }

            WebTTTFile.SheetName = WebTTTFile.SheetName.Replace("{anneeEC}", DateTime.Now.Year.ToString());

            if (Directory.Exists(pathSharepointSuiviSprint))
            {
                WebTTTFile.FullFileName = pathSharepointSuiviSprint + WebTTTFile.FolderName + WebTTTFile.FullFileName;
            }
            else if (Directory.Exists(pathSharepointSuiviSprint2))
            {
                WebTTTFile.FullFileName = pathSharepointSuiviSprint2 + WebTTTFile.FolderName + WebTTTFile.FullFileName;
            }
        }
    }
}
