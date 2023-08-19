using MajConsoSuiviSprint.Cli.Constants;
using MajConsoSuiviSprint.Cli.Model;
using System.Globalization;
using System.Text.RegularExpressions;

namespace MajConsoSuiviSprint.Cli.Utils
{
    internal static class InfoSprint
    {
        public static string GetFileNameSuiviSprintEC()
        {
            Console.WriteLine("GetFileNameSuiviSprintEC");
            int numSemaineEC = GetNumSemaine(DateTime.Now);
            string result = AppliConstant.TemplateSuiviDeSprint;
            string numSemaineFinsSprint;
            string numSemaineDebutSprint;
            int numPI = GetNumPIEC(numSemaineEC);

            if (numSemaineEC % 2 == 0)
            {
                numSemaineDebutSprint = (numSemaineEC - 1).ToString("D2");
                numSemaineFinsSprint = numSemaineEC.ToString("D2");
            }
            else
            {
                if (IsMondayToday())
                {
                    numSemaineEC -= 2;
                }
                numSemaineDebutSprint = numSemaineEC.ToString("D2");
                numSemaineFinsSprint = (numSemaineEC + 1).ToString("D2");
            }

            result = result
                    .Replace("{numDebSemaine}", numSemaineDebutSprint)
                    .Replace("{numFinDeSemaine}", numSemaineFinsSprint)
                    .Replace("{numPI}", numPI.ToString("D2"));

            return result;
        }

        public static bool IsActivityToManaged(string activite)
        {
            List<string> list = new() { AppliConstant.LblActiviteDev, AppliConstant.LblActiviteQual, AppliConstant.LblActiviteDev };
            return list.Contains(activite);
        }

        private static int GetNumPIEC(int numSemaine)
        {
            return (int)Math.Ceiling(numSemaine / 4.0);
        }

        public static int GetNumSemaine(DateTime date)
        {
            CultureInfo culture = new("fr-FR", false);
            Calendar calendar = culture.Calendar;
            CalendarWeekRule weekRule = culture.DateTimeFormat.CalendarWeekRule;
            DayOfWeek dayOfWeek = culture.DateTimeFormat.FirstDayOfWeek;

            int numberWeek = calendar.GetWeekOfYear(date, weekRule, dayOfWeek);
            return numberWeek;
        }

        private static bool IsMondayToday()
        {
            var jourCourant = DateTime.Now.DayOfWeek.ToString();

            return (jourCourant.Equals("Monday") || jourCourant.Equals("Lundi"));
        }

        public static (int debutSprint, int finSprint) ExtractSemainesSprint(string fichierDeSuivi)
        {
            var debutSemaineSprint = default(int);
            var finSemaineSprint = default(int);

            Match match = Regex.Match(fichierDeSuivi, @"CD13_PI\d{2}_S(\d{2})-(\d{2})");

            if (match.Success && match.Groups.Count > 2)
            {
                debutSemaineSprint = int.Parse(match.Groups[1].Value);
                finSemaineSprint = int.Parse(match.Groups[2].Value);
            }

            return (debutSemaineSprint, finSemaineSprint);
        }

        public static bool IsDemandeValid(ImportWebTTTExcelModel saisieInWebTTT, Dictionary<string, List<MaskSaisieModel>> maskAutorise)
        {
            bool result = false;

            if (saisieInWebTTT.Activite == AppliConstant.LblActiviteSpecification)
            {
                maskAutorise["Spec"].ForEach(mask =>
                {
                    if (saisieInWebTTT.NumeroDeDemande.Contains(mask.Rule))
                    {
                        result = true;
                    }
                });
            }
            else if (saisieInWebTTT.Activite.Equals(AppliConstant.LblActiviteDev) || saisieInWebTTT.Activite.Equals(AppliConstant.LblActiviteQual))
            {
                maskAutorise["DevQual"].ForEach(mask =>
                {
                    if (saisieInWebTTT.NumeroDeDemande.Contains(mask.Rule))
                    {
                        result = true;
                    }
                });

                if (result)
                {
                    result = !(saisieInWebTTT.Application == AppliConstant.LblApplicationParDefaut && !saisieInWebTTT.NumeroDeDemande.StartsWith("Inner"));
                }
            }

            return result;
        }

        public static string ModifyDemandeWebTTTToSuiviSprint(string numDemande, string application)
        {
            string result = default!;
            if (numDemande.ToUpper().StartsWith("INNERSOURCE") || numDemande.ToUpper().StartsWith("#"))
            {
                if (numDemande.ToUpper().StartsWith("INNERSOURCE"))
                {
                    result = result.Substring(12);
                }
                else if (numDemande.ToUpper().StartsWith("#"))
                {
                    result = result.Substring(1);
                }
                result = application + "-" + result;
            }
            else
            {
                string[] resultSplit = numDemande.Split('-');
                if (resultSplit.Length > 1)
                {
                    bool isnumeric = int.TryParse(resultSplit[1], out _);
                    if (isnumeric)
                    {
                        result = resultSplit[1];
                    }
                    else
                    {
                        isnumeric = int.TryParse(resultSplit[0], out _);
                        if (isnumeric)
                        {
                            result = resultSplit[0];
                        }
                        else
                        {
                            result = numDemande;
                        }
                    }
                }
                else
                {
                    result = numDemande.Replace("BUG", "").Replace("TS", "").Replace("US", "").Trim();
                }
            }

            return result.Trim();
        }
    }
}