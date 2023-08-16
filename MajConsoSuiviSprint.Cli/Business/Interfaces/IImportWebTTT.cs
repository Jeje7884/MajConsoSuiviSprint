using MajConsoSuiviSprint.Cli.Model;

namespace MajConsoSuiviSprint.Cli.Business.Interfaces
{
    internal interface IImportWebTTT
    {
        ResultImportWebTTT CheckSaisiesActiviteInWebTTT(List<ImportWebTTTExcelModel> allDataInWebTTT);
        void GenereExportCSVErreurSaisies(List<ErreurSaisieDemandeModel> erreursSaisiesDemandes);
    }
}