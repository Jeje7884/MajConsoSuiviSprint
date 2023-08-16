using MajConsoSuiviSprint.Cli.Constants;
using MajConsoSuiviSprint.Cli.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MajConsoSuiviSprint.Cli.Business
{
    internal class SuiviSprint
    {
        private static void AddDictionnaireTempsConsommeParDemandeEtParPole(Dictionary<string, TempsConsommeDemandeModel> tempsConsommesPardemandesEtParSprint, ImportWebTTTExcelModel dataWebTTT)
        {
            if (!tempsConsommesPardemandesEtParSprint.ContainsKey(dataWebTTT.NumeroDeDemande))
            {
                TempsConsommeDemandeModel tempsConsoDemande = new()
                {
                    //NumeroDeDemande = dataWebTTT.NumeroDeDemande,
                    Application = dataWebTTT.Application,
                    HeureTotaleDeDeveloppement = dataWebTTT.Activite.Equals(AppliConstant.LblActiviteDev) ? dataWebTTT.HeureDeclaree : 0f,
                    HeureTotaleDeQualification = dataWebTTT.Activite.Equals(AppliConstant.LblActiviteQual) ? dataWebTTT.HeureDeclaree : 0f
                };
                tempsConsommesPardemandesEtParSprint.Add(dataWebTTT.NumeroDeDemande, tempsConsoDemande);
            }
        }
    }
}
