namespace MajConsoSuiviSprint.Cli.Model
{
    public class WebTTTInfoModel
    {

        public int NbreSprintAPrendreEnCompte { get; set; } = 0;
        public string FolderName { get; set; } = default!;
        public string FileName { get; set; } = default!;
        public string SheetName { get; set; } = default!;
        public IReadOnlyCollection<HeadersWebTTT> Headers { get; set; } = new List<HeadersWebTTT>();


    }
}