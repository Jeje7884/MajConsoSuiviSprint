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
        
        public static bool IsPeriodeToManaged(int numSemaineSaisie, int numSemaineDebut, int numSemaineFin)
        {
            int numSemaineSaisie = GetNumSemaine(dateSaisie);
            
            return (numSemaineSaisie  >= numSemaineDebut) && (numSemaineSaisie <= numSemaineFin);
        }

        public static bool IsActivityToManaged(string activite)
        {
            List<string> list = new() { AppliConstant.LblActiviteDev,AppliConstant.LblActiviteQual,AppliConstant.LblActiviteDev};
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

        private static int[] ExtractSemainesSprintTableau(string fichierDeSuivi)
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
    }
}
