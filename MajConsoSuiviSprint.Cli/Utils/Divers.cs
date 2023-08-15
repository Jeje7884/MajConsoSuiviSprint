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
        public static void DisplayErrorMessageInConsole(string message)
        {
            Console.BackgroundColor = ConsoleColor.Yellow;
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"!!!!!! {message} !!!!!!!");
            Console.ResetColor();
        }

        public static void DisplayWarningMessageInConsole(string message)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($" --> {message} ");
            Console.ResetColor();
        }

        public static string GetFileNameFromFullPathFilename(string fullPath)
        {
            return Path.GetFileName(fullPath);
        }

        public static string GetPathFromFullPathFilename(string fullPath)
        {
            return Path.GetDirectoryName(fullPath) ?? "";
        }

        public static bool IsFileExist(string fullPath)
        {
            return File.Exists(fullPath);
        }

        public static bool IsFileWithPath(string path)
        {
            string fileName = Path.GetFileName(path);
            return path != fileName;

        }

        public static bool IsFileWithExtention(string file)
        {
            return Path.HasExtension(file);
        }

    }
}
