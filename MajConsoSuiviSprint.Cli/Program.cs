using MajConsoSuiviSprint.Cli.Business;
using MajConsoSuiviSprint.Cli.Constants;
using MajConsoSuiviSprint.Cli.Utils;

namespace MajConsoSuiviSprint.Cli
{
    public static class Programm
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("*****************************************"); ;
            Console.WriteLine($"Début du script à {DateTime.Now}");
            Console.WriteLine("*****************************************");

            try
            {
                string pathConfigJson = InitPathJsonConfig(args);

                ConfigurationsApp configurationProcess = new(pathConfigJson);

                var importWebTTT = new ImportWebTTT(configurationProcess);
                var result = importWebTTT.ImportInfosFromWebTTT();

                if (configurationProcess.WebTTTInfoConfig.FullFileName is null)
                {
                    throw new Exception("Les paramétrages en lien avec le fichier WebTTT sont erronés");
                }

            }
            catch (Exception ex)
            {
                Divers.DisplayErrorMessageInConsole(ex.Message);
            }
            finally
            {
                Console.WriteLine("Fin du script. Merci d'avoir joué");
                Console.WriteLine("Au revoir");
                Console.WriteLine("Taper sur 'Entrée' pour terminer...");
                Console.ReadLine();
            }
        }

        private static string InitPathJsonConfig(string[] args)
        {
            string pathConfigJson = default!;
            if (args.Length.Equals(0) || args.Length.Equals(1))
            {
                if (args.Length == 1)
                {
                    pathConfigJson = args[0];
                }
                else
                {
                    Console.WriteLine(" - Tapez \"Entrer\" pour choisir le fichier de config par défaut  ");
                    Console.WriteLine(" - sinon saisir le fichier de config à utiliser  : ");
                    Console.WriteLine("      => (par exemple c:\\temp\\MonFichierAppSettings.json) ");
                    Console.WriteLine("       ou seulement le nom du fichier si dans le répertoire courant (par exemple MonFichierAppSettings.json) ");

                    string? choix = Console.ReadLine() ?? "";

                    if (string.IsNullOrEmpty(choix))
                    {

                        pathConfigJson = Directory.GetCurrentDirectory() + "\\" + AppliConstant.FileAppSettingsParDefaut;
          
    }
                    else 
                    {
                        if (Divers.IsFileWithPath(choix))
                        {
                            pathConfigJson = choix;
                        }
                        else
                        {
                            if(!Divers.IsFileWithExtention(choix))
                            {
                                choix +=  AppliConstant.ExtensionAppSettings;
                            }
                            pathConfigJson = Directory.GetCurrentDirectory() + "\\" + choix;
                        }
                        
                    }
                }
                if (!Divers.IsFileExist(pathConfigJson))
                {
                    throw new Exception($"Le fichier json passé {pathConfigJson}");
                }
            }
            else if (args.Length > 1)
            {
                throw new Exception("Erreur dans le nombre de paramètre passé");
            }

            return pathConfigJson;
        }
    }
}

