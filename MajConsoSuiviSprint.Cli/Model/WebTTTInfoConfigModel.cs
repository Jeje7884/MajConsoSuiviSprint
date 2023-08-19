using MajConsoSuiviSprint.Cli.Utils;

namespace MajConsoSuiviSprint.Cli.Model
{
    public class WebTTTInfoConfigModel
    {
        public string FolderName { get; set; } = default!;
        public string FileName { get; set; } = default!;
        public string FullFileName { get; set; } = default!;
        public string SheetName { get; set; } = default!;  

        public IReadOnlyCollection<ListConfigsWebTTTModel> JoursFeries { get; set; } = default!;
    
        public string Path { get; set; } = default!;
        public string FileBilansErreurFormatSaisieDemande { get; set; } = default!;
        public string FileBilansErreurTempsSaisieSemaine { get; set; } = default!;
        public int NbreDeSemaineAPrendreAvtLaSemaineEnCours { get; set; } = 0;

        public bool TopLaunchBilans { get; set; } = default!;

        public int NumeroDeSemaineAPartirDuquelChecker
        {
            get { return CalculDuNumeroDeSemaineDePriseEnCompte(); }
        }

        private int CalculDuNumeroDeSemaineDePriseEnCompte()
        {
            return (InfoSprint.GetNumSemaine(DateTime.Now) - NbreDeSemaineAPrendreAvtLaSemaineEnCours);
        }

        public int NbreJourSemaine { get; set; }
        public int NbreHeureTotaleActiviteJour { get; set; }

        public Dictionary<string, List<MaskSaisieModel>> ReglesSaisiesAutorisesParActivite { get; set; } = new Dictionary<string, List<MaskSaisieModel>>();
    }

    public class ListConfigsWebTTTModel
    {
        public string Value { get; set; } = default!;
    }

    public class MaskSaisieModel
    {
        public string Rule { get; set; } = default!;
    }
}