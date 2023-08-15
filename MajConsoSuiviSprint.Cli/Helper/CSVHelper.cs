using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using CsvHelper;
using System.Threading.Tasks;
using MajConsoSuiviSprint.Cli.Utils;

namespace MajConsoSuiviSprint.Cli.Helper
{
    internal class CSVHelper
    {
        public static void GenerateCSVFile(string fileName, List<CleValeur<string, object>> data, bool isAppend, string typeEncoding = "utf8BOM")
        {
            Console.WriteLine("CSVModule.GenerateCSVFile");
         
            using (var writer = new StreamWriter("output.csv", isAppend, GetEncoding(typeEncoding)))
            {
               
                using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                {
                    csv.WriteRecords(data);
                }
            }

           
        }

        public static void GenerateCSVFile(string fileName, List<string> data, bool isAppend, string typeEncoding = "utf8BOM")
        {
            Console.WriteLine("CSVModule.GenerateCSVFile");

            using (var writer = new StreamWriter("output.csv", isAppend, GetEncoding(typeEncoding)))
            {

                using var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);
             
                csv.WriteRecords(data);
            }     
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
                    throw new NotSupportedException($"Encoding '{typeEncoding}' is not supported.");
            }
        }

      
    }
}
