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
    } 
}
