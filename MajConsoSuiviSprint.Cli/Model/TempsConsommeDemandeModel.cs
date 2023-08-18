namespace MajConsoSuiviSprint.Cli.Model
{
    internal class TempsConsommeDemandeModel
    {
        public string Application { get; set; } = default!;
        public string NumeroDeDemande { get; set; } = default!;
        public float HeureTotaleDeDeveloppement { get; set; } = default!;
        public float HeureTotaleDeQualification { get; set; } = default!;
        public bool IsDemandeValide { get; set; } = default!;

    }
}