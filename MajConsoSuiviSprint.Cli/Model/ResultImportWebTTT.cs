namespace MajConsoSuiviSprint.Cli.Model
{

    internal class ResultImportWebTTT
    {
        public Dictionary<string, ErreurSaisieDemandeModel> ErreursSaisiesDemandes { get; set; } = default!;
        public Dictionary<string, TempsConsommeDemandeModel> TempsConsommesPardemandesEtParSprint { get; set; } = default!;
        //public Dictionary<string, Dictionary<int,TempsDeclareCollabModel>> TempsDeclaresTotalSemaineParCollab { get; set; } = default!;
        //public Dictionary<string, Dictionary<int,float>> TempsDeclaresTotalSemaineParCollab { get; set; } = default!;
        public List<ErreurSaisieRemplissageTempsParSemaineModel> ErreursSaisiesRemplissageTempsCollabParSemaine{ get; set; } = default!;
    }
}
