namespace MajConsoSuiviSprint.Cli.Model
{
    public class WebTTTInfoModel
    {
        public string Path { get; set; } = default!;
        public string FileBilanErreurCSV { get; set; } = default!;
        public int NbreSprintAPrendreEnCompte { get; set; } = 0;

        public string FileName { get; set; } = default!;
        public string SheetName { get; set; } = default!;
        public string FullFileName { get; set; } = default!;
        public  int NumDebutSemaineAImporter { get; set; }
        public int NumFinSemaineAImporter { get; set; }
        public IReadOnlyCollection<HeadersWebTTTModel> Headers { get; set; } = default!;

        public List<MaskSaisieModel> ReglesSaisiesAutorisePourUneActivite { get; set; } = default!;

        public Dictionary<string, List<MaskSaisieModel>> ReglesSaisiesAutorisesParActivite { get; set; } = new Dictionary<string, List<MaskSaisieModel>>();
    }

}