namespace MajConsoSuiviSprint.Cli.Model
{
    internal class ErreurSaisieDemandeModel
    {
        public string Application { get; set; } = default!;
        public string Activite { get; set; } = default!;
        public string NumeroDemande { get; set; } = default!;
        public string Qui { get; set; } = default!;
        public string DateDeSaisie { get; set; } = default!;
        public string DetailErreur { get; set; } = default!;

    }
}
