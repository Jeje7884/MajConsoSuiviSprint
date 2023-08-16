namespace MajConsoSuiviSprint.Cli.Model
{
    internal class ResultImportWebTTT
    {
        public Dictionary<string, ErreurSaisieDemandeModel> ErreursSaisiesDemandes { get; set; } = default!;
        public Dictionary<string, TempsConsommeDemandeModel> TempsConsommesPardemandesEtParSprint { get; set; } = default!;
        public List<SaisieRemplissageTempsParSemaineModel> ErreursSaisiesRemplissageTempsCollabParSemaine { get; set; } = default!;
    }
}