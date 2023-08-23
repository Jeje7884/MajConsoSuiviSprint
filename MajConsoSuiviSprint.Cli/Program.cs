using MajConsoSuiviSprint.Cli.Business;
using MajConsoSuiviSprint.Cli.Constants;
using MajConsoSuiviSprint.Cli.Model;
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
                if (args.Length > 1)
                {
                    throw new Exception("Erreur dans le nombre de paramètre passé");
                }

                string pathConfigJson = InitPathJsonConfig(args);

                ConfigurationsApp configurationProcess = new(pathConfigJson);

                if (configurationProcess.WebTTTInfoConfig.FullFileName is null)
                {
                    throw new Exception("Les paramétrages \"Path\"  et  \"FileName\" de la saction WebTTT sont erronés");
                }

                if (configurationProcess.SuiviSprintInfoConfig.FullFileName is null)
                {
                    throw new Exception("Les paramétrages \"Path\"  et  \"FileName\" de la saction SuiviSprint sont erronés");
                }

                var importWebTTT = new ImportWebTTT(configurationProcess);

                IList<ImportWebTTTExcelModel> resultImport = importWebTTT.ImportInfosFromWebTTT();
                ResultImportWebTTT resultAnalyseImport = importWebTTT.CheckSaisiesActiviteInWebTTT(resultImport);

                importWebTTT.GenereExportCSVErreurSaisies(resultAnalyseImport.ErreursSaisiesDemandes);
                importWebTTT.GenereBilanErreurTempsConsommeSemaine(resultAnalyseImport.SaisiesRemplissageTempsCollabParSemaine);
            }
            catch (Exception ex)
            {
                Tools.DisplayErrorMessageInConsole(ex.Message);
            }
            finally
            {
                Console.WriteLine("Fin du script. Merci d'avoir joué");
                Console.WriteLine("Au revoir");
                Console.WriteLine("Taper sur 'Entrée' pour terminer...");
                Console.ReadLine();
            }
        }

        private static string InitPathJsonConfig(string[] @params)
        {
            string pathConfigJson;

            if (@params.Length == 1)
            {
                pathConfigJson = @params[0];
                if (!Tools.IsFileWithPath(pathConfigJson))
                {
                    pathConfigJson = Directory.GetCurrentDirectory() + "\\" + pathConfigJson;
                }
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
                    if (Tools.IsFileWithPath(choix))
                    {
                        pathConfigJson = choix;
                    }
                    else
                    {
                        if (!Tools.IsFileWithExtention(choix))
                        {
                            choix += AppliConstant.ExtensionAppSettings;
                        }
                        pathConfigJson = Directory.GetCurrentDirectory() + "\\" + choix;
                    }
                }
            }
            if (!Tools.IsFileExist(pathConfigJson))
            {
                throw new Exception($"Le fichier json {pathConfigJson} n'existe pas");
            }

            return pathConfigJson;
        }
    }
}