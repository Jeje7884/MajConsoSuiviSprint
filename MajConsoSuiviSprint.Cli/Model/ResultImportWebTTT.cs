namespace MajConsoSuiviSprint.Cli.Model
{
    internal class ResultImportWebTTT
    {
        public List<ErreurSaisieFormatDemandeModel> ErreursSaisiesDemandes { get; set; } = default!;

        public List<SaisieRemplissageTempsCollabParSemaineModel> SaisiesRemplissageTempsCollabParSemaine { get; set; } = default!;
    }
}