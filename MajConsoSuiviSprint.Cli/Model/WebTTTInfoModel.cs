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

    }
    public class HeadersWebTTT

    {
        public string Valeur { get; set; } = default!;
    }
}