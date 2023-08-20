using MajConsoSuiviSprint.Cli.Model;
using Newtonsoft.Json;
using System.Diagnostics;

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

        public static void DisplayInfoMessageInConsole(string message)
        {
            Console.ForegroundColor = ConsoleColor.Blue;
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

        public static void DeleteFile(string fullPath)
        {
            if (!IsFileOpened(fullPath))
            {
                File.Delete(fullPath);
            }
            else
            {
                throw new Exception($"Le fichier {fullPath} ne peut être supprimé car il est ouvert");
            }
        }

        public static void LaunchProcess(string fullPath)
        {
            Process.Start("explorer.exe", fullPath);


        }

        public static void testPwh()
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>
            {
                { "Param1", "Value1" },
                { "Param2", "Value2" },
                { "Param3", "Value3" }
            };

            //Dictionary<string, TempsConsommeDemandeModel> parameters = new()
            //{
            //    {
            //        "2022",new TempsConsommeDemandeModel{
            //            Application ="SRS",
            //            HeureTotaleDeDeveloppement =10,
            //            HeureTotaleDeQualification =10,
            //            NumeroDeDemande ="TS-2022"
            //        }


            //    },
            //    {
            //        "2025",new TempsConsommeDemandeModel{
            //            Application ="SRS",
            //            HeureTotaleDeDeveloppement =10,
            //            HeureTotaleDeQualification =10,
            //            NumeroDeDemande ="TS-2025"

            //        }
            //    }

            //};

            // Convertissez le dictionnaire en une chaîne JSON pour le passer au script PowerShell
            //string jsonParameters = Newtonsoft.Json.JsonConvert.SerializeObject(parameters);
            //String jsonParameters = JsonConvert.SerializeObject(parameters, Newtonsoft.Json.Formatting.None);
            String jsonParameters = MyDictionaryToJsonTest2(parameters);
            //String jsonParameters = JsonConvert.SerializeObject(jsonParametersInit, Newtonsoft.Json.Formatting.None);
            // Chemin vers le script PowerShell à exécuter
            string scriptPath = @".\Utils\Test.ps1";

            // Commande à exécuter dans PowerShell avec les paramètres JSON
            string powershellCommand = $"-ExecutionPolicy Bypass -File \"{scriptPath}\"  '{jsonParameters}'";

            // Exécution du processus PowerShell
            ProcessStartInfo psi = new ProcessStartInfo("powershell.exe", powershellCommand)
            {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            Process process = new Process { StartInfo = psi };
            process.Start();

            // Capture de la sortie standard et des éventuelles erreurs
            string output = process.StandardOutput.ReadToEnd();
            string errors = process.StandardError.ReadToEnd();

            process.WaitForExit();

            Console.WriteLine("Output:");
            Console.WriteLine(output);

            if (!string.IsNullOrEmpty(errors))
            {
                Console.WriteLine("Errors:");
                Console.WriteLine(errors);
            }
        }

        private static string MyDictionaryToJsonTest(Dictionary<string, string> dict)
        {
            var entries = dict.Select(d =>
                string.Format("\"\"{0}\":{1}", d.Key, string.Join(",", d.Value)));
            return "{" + string.Join(",", entries) + "}";
        }

        private static string MyDictionaryToJsonTest2(Dictionary<string, string> dict)
        {
            string result=default!;
            foreach (var entry in dict)
            {
                result += $"\"\"{entry.Key}\"\":{entry.Value},";
                //result += @"""" + result + @"""" + entry.Key + ":" + entry.Value + ",";

            }
            result = "{" + result +"}";
            //var entries = dict.Select(d =>
            //    string.Format("\"\"{0}\":{1}", d.Key, string.Join(",", d.Value)));
            //return "{" + string.Join(",", entries) + "}";
            return result;
        }
        private static string MyDictionaryToJson(Dictionary<int, List<int>> dict)
        {
            var entries = dict.Select(d =>
                string.Format("\"{0}\": [{1}]", d.Key, string.Join(",", d.Value)));
            return "{" + string.Join(",", entries) + "}";
        }
    }
    
}