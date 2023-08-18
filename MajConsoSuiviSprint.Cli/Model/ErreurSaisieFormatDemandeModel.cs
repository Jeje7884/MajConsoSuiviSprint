namespace MajConsoSuiviSprint.Cli.Model
{
    internal class ErreurSaisieFormatDemandeModel
    {
        public string Qui { get; set; } = default!;
        public string Application { get; set; } = default!;
        public string Activite { get; set; } = default!;
        public string NumeroDeDemande { get; set; } = default!;
        
        public string DateDeSaisie { get; set; } = default!;
        public string DetailErreur { get; set; } = default!;
    }
}