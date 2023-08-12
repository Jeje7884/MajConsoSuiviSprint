namespace MajConsoSuiviSprint.Cli.Utils
{
    internal static class Divers
    {
        public static bool IsFileOpened(string fullNameFile)
        {
            try
            {

                File.OpenRead(fullNameFile).Close();
                return false;
            }
            catch
            {
                return true;
            }
        }
        public static void DisplayErrorInConsole(string message)
        {
            Console.BackgroundColor = ConsoleColor.Yellow;
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"!!!!!! {message} !!!!!!!");
            Console.ResetColor();
        }
    }
}
