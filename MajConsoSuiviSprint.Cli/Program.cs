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
                string pathConfigJson = string.Empty;
                if (args.Length.Equals(1))

                {
                    pathConfigJson = args[0];
                }
                else if (args.Length > 1)
                {
                    throw new Exception("Erreur dans le nombre de paramètre passé");
                }
                Configuration configurationProcess = new(pathConfigJson);

                var result = InfoSprint.GetFileNameSuiviSprintEC();
                Console.WriteLine("La valeur du fichier de siuvi est en cours est " + result);
                if (configurationProcess.WebTTTModel.FullFileName is null)
                {
                    throw new Exception("Les paramétrages en lien avec le fichier WebTTT sont erronés");
                }

            }
            catch (Exception ex)
            {
                Divers.DisplayErrorInConsole(ex.Message);
            }
            finally
            {
                Console.WriteLine("Fin du script. Merci d'avoir joué");
                Console.WriteLine("Au revoir");
                Console.WriteLine("Taper sur 'Entrée' pour terminer...");
                Console.ReadLine();
            }
        }
    }
}

