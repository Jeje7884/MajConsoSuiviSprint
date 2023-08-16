using CsvHelper;
using System.Globalization;
using System.Text;

namespace MajConsoSuiviSprint.Cli.Helper
{
    internal class CSVHelper
    {
        //public static void GenerateCSVFile(string fileName, List<CleValeur<string, object>> data, bool isAppend, string typeEncoding = "utf8BOM")
        //{
        //    Console.WriteLine("CSVModule.GenerateCSVFile");

        //    using var writer = new StreamWriter(fileName, isAppend, GetEncoding(typeEncoding));

        //    using var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);
        //    csv.WriteRecords(data);
        //}

        public static void GenerateCSVFile<T>(string fileName, List<T> data, bool isAppend, string typeEncoding = "utf8BOM")
        {
            Console.WriteLine("CSVModule.GenerateCSVFile");

            using var writer = new StreamWriter(fileName, isAppend, GetEncoding(typeEncoding));

            using var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);
            csv.WriteRecords(data);
        }
        

        private static Encoding GetEncoding(string typeEncoding)
        {
            switch (typeEncoding)
            {
                case "utf8BOM":
                    return new UTF8Encoding(true);

                case "utf8":
                    return new UTF8Encoding(false);

                default:
                    throw new NotSupportedException($"Encoding '{typeEncoding}' n'est pas géré dans l'export CSV");
            }
        }
    }
}