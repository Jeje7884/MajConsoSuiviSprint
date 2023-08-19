using CsvHelper;
using CsvHelper.Configuration;
using System.Globalization;
using System.Text;

namespace MajConsoSuiviSprint.Cli.Helper
{
    internal class CSVHelper
    {
        public static void GenerateCSVFile<T>(string fileName, IList<T> data, bool isAppend, string typeEncoding = "utf8BOM")
        {
            Console.WriteLine("CSVModule.GenerateCSVFile");
            CsvConfiguration delimiter = GetDelimiter();
            using var writer = new StreamWriter(fileName, isAppend, GetEncoding(typeEncoding));
            using var csv = new CsvWriter(writer, delimiter);
            csv.WriteRecords(data);
        }

        private static CsvConfiguration GetDelimiter()
        {
            return new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = ";"
            };
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