using MajConsoSuiviSprint.Cli.Business;
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
                string? pathConfigJson = default!;

                pathConfigJson = InitPathJsonConfig(args, pathConfigJson);

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

        private static string InitPathJsonConfig(string[] args, string pathConfigJson)
        {
            if (args.Length.Equals(0) || args.Length.Equals(1))
            {
                string? choix = string.Empty;
                const string valParDefaut = "D";
                if (args.Length == 1)
                {
                    pathConfigJson = args[0];
                }
                else
                {
                    Console.WriteLine(" - Tapez \"D\" ((ou \"Entrer\") pour choisir le fichier de config par défaut  ");
                    Console.WriteLine(" - sinon saisir le fichier de config à utiliser (par exemple c:\\temp\\MonFichierAppSettings.json) : ");

                    choix = Console.ReadLine() ?? valParDefaut;

                    if (string.IsNullOrEmpty(choix))
                    {
                        choix = valParDefaut;
                    }
                    if (!choix.Equals(valParDefaut))
                    {
                        pathConfigJson = choix;
                    }
                }
                if (!choix.Equals(valParDefaut) && !Divers.IsFileExist(pathConfigJson))
                {
                    throw new Exception("Le fichier json passé en paramètre n'existe pas");
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