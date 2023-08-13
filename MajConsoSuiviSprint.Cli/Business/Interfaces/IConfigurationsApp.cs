using MajConsoSuiviSprint.Cli.Model;

namespace MajConsoSuiviSprint.Cli.Business.Interfaces
{
    internal interface IConfigurationsApp
    {
        public WebTTTInfoModel WebTTTModel { get; }
        public SuiviSprintModel SuiviSprintModel { get; }
    }
}
