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

        public IReadOnlyCollection<HeadersWebTTT> Headers { get; set; } = default!;

        public List<MaskSaisie> ReglesSaisiesAutorisePourUneActivite { get; set; } = default!;

        public Dictionary<string, List<MaskSaisie>> ReglesSaisiesAutorisesParActivite { get; set; } = new Dictionary<string, List<MaskSaisie>>();
    }
    public class HeadersWebTTT

    {
        public string Value { get; set; } = default!;
    }

    public class MaskSaisie
    {
        public string Rule { get; set; } = default!;
    }
}