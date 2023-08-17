namespace MajConsoSuiviSprint.Cli.Model
{
    internal class ResultImportWebTTT
    {
        public List<ErreurSaisieDemandeModel> ErreursSaisiesDemandes { get; set; } = default!;
        //public Dictionary<string, ErreurSaisieDemandeModel> ErreursSaisiesDemandes { get; set; } = default!;
        //public Dictionary<string, TempsConsommeDemandeModel> TempsConsommesPardemandesEtParSprint { get; set; } = default!;
        public List<SaisieRemplissageTempsCollabParSemaineModel> SaisiesRemplissageTempsCollabParSemaine { get; set; } = default!;
    }
}