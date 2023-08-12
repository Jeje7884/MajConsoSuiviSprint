using MajConsoSuiviSprint.Cli.Constants;
using System.Globalization;
using System.Text.RegularExpressions;

namespace MajConsoSuiviSprint.Cli.Utils
{
    internal static class InfoSprint
    {
        public static string GetFileNameSuiviSprintEC()

        {
            Console.WriteLine("GetFileNameSuiviSprintEC");
            int numSemaineEC = GetNumSemaineEC();
            string result = SprintConstant.templateSuiviDeSprint;
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
                    numSemaineEC = -2;
                }
                numSemaineDebutSprint = numSemaineEC.ToString("D2");
                numSemaineFinsSprint = (numSemaineEC + 1).ToString("D2");
            }

            result = result
                    .Replace("{numDebSemaine}", numSemaineDebutSprint)
                    .Replace("{numFinDeSemaine}", numSemaineFinsSprint)
                    .Replace("{numPI}", numPI.ToString("D2"));

            var semaines=ExtractSemainesSprint(result);
            Console.WriteLine($"Semaine debut de sprint {semaines[0]}");
            Console.WriteLine($"Semaine fin de sprint {semaines[1]}");
            return result;
        }
        private static int GetNumPIEC(int numSemaine)
        {
            return (int)Math.Ceiling(numSemaine / 4.0);
        }

        private static int GetNumSemaineEC()
        {

            DateTime date = DateTime.Now;

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

        private static int[] ExtractSemainesSprint(string fichierDeSuivi)
        {
            var debutSemaineSprint = default(int);
            var finSemaineSprint = default(int);
            int[] resultSemainesSprint = new int[2];

            Match match = Regex.Match(fichierDeSuivi, @"CD13_PI\d{2}_S(\d+)-(\d+)");

            if (match.Success && match.Groups.Count > 2)
            {
                debutSemaineSprint = int.Parse(match.Groups[1].Value);
                finSemaineSprint = int.Parse(match.Groups[2].Value);
            }
            resultSemainesSprint[0] = debutSemaineSprint;
            resultSemainesSprint[1] = finSemaineSprint;
            return resultSemainesSprint;
        }
    }
}
