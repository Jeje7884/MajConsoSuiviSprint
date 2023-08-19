using MajConsoSuiviSprint.Cli.Model;

namespace MajConsoSuiviSprint.Cli.Business.Interfaces
{
    internal interface IImportWebTTT
    {
        ResultImportWebTTT CheckSaisiesActiviteInWebTTT(IList<ImportWebTTTExcelModel> allDataInWebTTT);

        void GenereExportCSVErreurSaisies(IList<ErreurSaisieFormatDemandeModel> erreursSaisiesDemandes);
    }
}