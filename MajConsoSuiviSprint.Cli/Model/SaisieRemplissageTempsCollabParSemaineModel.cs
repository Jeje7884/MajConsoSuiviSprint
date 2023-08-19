using CsvHelper.Configuration.Attributes;

namespace MajConsoSuiviSprint.Cli.Model
{
    internal class SaisieRemplissageTempsCollabParSemaineModel
    {
        public string Qui { get; set; } = default!;
        public int NumeroDeSemaine { get; set; } = default!;
        public float TotalHeureDeclaree { get; set; } = default!;

        [Ignore]
        public int TotalHeureTempsPlein { get; set; } = default!;
    }
}