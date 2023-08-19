namespace MajConsoSuiviSprint.Cli.Model
{
    internal class ImportWebTTTExcelModel
    {
        public string Activite { get; set; } = default!;
        public string TrigrammeCollab { get; set; } = default!;
        public string LblDateDeSaisie { get; set; } = default!;
        public DateTime DateDeSaisie { get; set; } = default!;
        public float HeureDeclaree { get; set; } = default!;

        public string Application { get; set; } = default!;
        public string NumeroDeDemande { get; set; } = default!;
        public int NumeroDeSemaineDateActivite { get; set; } = default!;
    }
}