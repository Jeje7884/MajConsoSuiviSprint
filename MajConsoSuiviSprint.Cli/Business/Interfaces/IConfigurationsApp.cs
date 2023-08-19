using MajConsoSuiviSprint.Cli.Model;

namespace MajConsoSuiviSprint.Cli.Business.Interfaces
{
    internal interface IConfigurationsApp
    {
        public WebTTTInfoConfigModel WebTTTInfoConfig { get; }
        public SuiviSprintInfoConfigModel SuiviSprintInfoConfig { get; }
    }
}