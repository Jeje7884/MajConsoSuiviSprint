using MajConsoSuiviSprint.Cli.Model;
using MajConsoSuiviSprint.Cli.Utils;
using Microsoft.Extensions.Configuration;

namespace MajConsoSuiviSprint.Cli.Business
{
    internal class Configuration
    {
        private const string WebTTTSection = "WebTTTFile";

        public string FileBilanErreurCSV { get; set; }
        public string PathSharepointSuiviSprint { get; set; }
        public string PathSharepointSuiviSprint2 { get; set; }
        public WebTTTInfoModel WebTTTModel { get; set; } = default!;

        public Configuration()
        {
            PathSharepointSuiviSprint = "";

            IConfiguration config = new ConfigurationBuilder()
            .AddJsonFile("appSettings.json", optional: false, reloadOnChange: true)
            .Build();

            //FileBilanErreurCSV = config.GetSection(nameof(Applications)).Get<List<Application>>().AsReadOnly();
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


            InitWebTTT(ref webTTTInfo, PathSharepointSuiviSprint, PathSharepointSuiviSprint);
            WebTTTModel = webTTTInfo;

        }
        private static string InitFileNameBilanErreurSaisieDansWebTTT()
        {
            return "";
        }

        private static void InitWebTTT(ref WebTTTInfoModel WebTTTFile, string pathSharepointSuiviSprint, string pathSharepointSuiviSprint2)
        {
            if (string.IsNullOrEmpty(WebTTTFile.FileName))
            {
                WebTTTFile.FileName = InfoSprint.GetFileNameSuiviSprintEC();
            }
            else
            {
                WebTTTFile.FileName = WebTTTFile.FileName.Replace("{anneeEC}", DateTime.Now.Year.ToString());
            }

            WebTTTFile.SheetName = WebTTTFile.FileName.Replace("{anneeEC}", DateTime.Now.Year.ToString());

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
