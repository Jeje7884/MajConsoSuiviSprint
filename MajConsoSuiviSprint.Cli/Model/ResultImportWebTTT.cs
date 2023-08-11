namespace MajConsoSuiviSprint.Cli.Model
{

    internal class ResultImportWebTTT
    {
        public Dictionary<string, ErreurSaisieDemandeModel> ErreursSaisiesDemandes { get; set; } = default!;
        public Dictionary<string, TempsConsommesDemandesModel> TempsConsommesPardemandesEtParSprint { get; set; } = default!;
        public Dictionary<string, ErreurSaisieTempsDeclareCollabModel> ErreursSaisiesTempsDeclarseCollab { get; set; } = default!;
    }
}
