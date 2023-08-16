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
            return typeEncoding switch
            {
                "utf8BOM" => new UTF8Encoding(true),
                "utf8" => new UTF8Encoding(false),
                _ => throw new NotSupportedException($"Encoding '{typeEncoding}' n'est pas géré dans l'export CSV"),
            };
        }
    }
}