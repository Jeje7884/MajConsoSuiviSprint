using System.Globalization;

namespace MajConsoSuiviSprint.Cli.Helper
{
    internal static class InfoSprint
    {
        public static string GetFileNameSuiviSprintEC()

        {
            var numSemaineEC = GetNumSemaineEC();
            Console.WriteLine($"Le numéro de semaine est {numSemaineEC}");
            return "";
        }
        private static string GetNumPIEC()
        {
            return "";
        }

        private static int GetNumSemaineEC()
        {

            DateTime date = DateTime.Now; // Remplacez par votre date

            CultureInfo culture = new("fr-FR", false);
            Calendar calendar = culture.Calendar;
            CalendarWeekRule weekRule = culture.DateTimeFormat.CalendarWeekRule;
            DayOfWeek firstDayOfWeek = culture.DateTimeFormat.FirstDayOfWeek;

            int weekNumber = calendar.GetWeekOfYear(date, weekRule, firstDayOfWeek);
            return weekNumber;
            //Console.WriteLine($"Numéro de semaine : {weekNumber}");
        }
    }
}
