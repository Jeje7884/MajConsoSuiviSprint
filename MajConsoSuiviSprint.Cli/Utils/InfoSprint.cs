using MajConsoSuiviSprint.Cli.Constants;
using System.Globalization;

namespace MajConsoSuiviSprint.Cli.Utils
{
    internal static class InfoSprint
    {
        public static string GetFileNameSuiviSprintEC()

        {
            int numSemaineEC = GetNumSemaineEC();
            string result = SprintConstant.templateSuiviDeSprint;
            string numSemaineFinsSprint;
            string numSemaineDebutSprint;
            int numPI = GetNumPIEC(numSemaineEC);
            
            if (numSemaineEC % 2 == 0)
            {
                numSemaineDebutSprint = (numSemaineEC - 1).ToString("D2");
                numSemaineFinsSprint = (numSemaineEC).ToString("D2");
            }
            else
            {
                if (IsMondayToday())
                {
                    numSemaineEC = -2;
                }
                numSemaineDebutSprint = (numSemaineEC).ToString();
                numSemaineFinsSprint = (numSemaineEC + 1).ToString();
            }

            result = result
                    .Replace("{numDebSemaine}", numSemaineDebutSprint)
                    .Replace("{numFinDeSemaine}", numSemaineFinsSprint)
                    .Replace("{numPI}", numPI.ToString("D2"));

            Console.WriteLine($"résulat {result}");
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
    }
}
