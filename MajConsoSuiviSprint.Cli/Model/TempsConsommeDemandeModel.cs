namespace MajConsoSuiviSprint.Cli.Model
{
    internal class TempsConsommeDemandeModel
    {
        public string Application { get; set; } = default!;
        //public string NumeroDeDemande { get; set; } = default!;
        public float HeureDeDeveloppement { get; set; } = default!;
        public float HeureDeQualification { get; set; } = default!;
        public bool IsDemandeValide { get; set; } = default!;
    }
}