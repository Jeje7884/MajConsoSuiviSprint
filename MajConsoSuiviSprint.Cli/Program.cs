using MajConsoSuiviSprint.Cli.Helper;

namespace MajConsoSuiviSprint.Cli
{
    public static class Programm
    {
        public static void Main(string[] args)
        {

            Console.WriteLine("*****************************************"); ;
            Console.WriteLine("Lancement du script");
            Console.WriteLine("*****************************************");


            try
            {
                //TODO
                var result = InfoSprint.GetFileNameSuiviSprintEC();
                //var test = InfoSprint.GetFileNameSuiviSprintEC();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Console.WriteLine("Fin du script");
                Console.WriteLine("Tapper sur 'Entrée' pour terminer...");
                Console.ReadLine();

            }

        }
    }

}
