using MajConsoSuiviSprint.Cli.Model;
using Microsoft.Extensions.Configuration;

namespace MajConsoSuiviSprint.Cli.Business
{
    internal class Configuration
    {
        private const string WebTTTSection = "WebTTTFile";

        public string FileBilanErreurCSV { get; set; }
        public string PathSharepointSuiviSprint { get; set; }
        public WebTTTInfoModel WebTTTFile { get; set; } = default!;

        public Configuration()
        {
            PathSharepointSuiviSprint = "";

            IConfiguration config = new ConfigurationBuilder()
            .AddJsonFile("appSettings.json", optional: false, reloadOnChange: true)
            .Build();

            //FileBilanErreurCSV = config.GetSection(nameof(Applications)).Get<List<Application>>().AsReadOnly();
            FileBilanErreurCSV = config.GetValue<string>("fileNameBilanErreurSaisieDansWebTTT") ?? "";
            PathSharepointSuiviSprint = config.GetValue<string>(nameof(PathSharepointSuiviSprint)) ?? "";
            
            var sectionWebTTT = config.GetSection(WebTTTSection);

            WebTTTFile.FolderName = sectionWebTTT.GetValue<string>("FolderName") ?? "";
            WebTTTFile.FileName = sectionWebTTT.GetValue<string>("FileName") ?? "";
            WebTTTFile.SheetName = sectionWebTTT.GetValue<string>("SheetName") ?? "";
            WebTTTFile.NbreSprintAPrendreEnCompte = sectionWebTTT.GetValue<int>("NbreSprintAPrendreEnCompte");
            WebTTTFile.Headers = sectionWebTTT
                                .GetSection("Headers")
                                .Get<List<HeadersWebTTT>>()
                                ?.AsReadOnly()
                                ?? new List<HeadersWebTTT>().AsReadOnly();


        }
        private static string InitFileNameBilanErreurSaisieDansWebTTT()
        {
            return "";
        }
    }
}
