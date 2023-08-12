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
                //TODO
                var result = InfoSprint.GetFileNameSuiviSprintEC();

                Console.WriteLine(result);
                Configuration configurationProcess = new();
                if (configurationProcess.WebTTTModel.FullFileName is null)
                {
                    throw new Exception("Les paramétrages en lien avec le fichier WebTTT sont erronés");
                }
                var date = DateTime.Now.Year;

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

