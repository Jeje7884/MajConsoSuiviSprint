using MajConsoSuiviSprint.Cli.Utils;

namespace MajConsoSuiviSprint.Cli.Model
{
    public class SuiviSprintInfoConfigModel
    {
     
        private string? _fileName;
        public string Path { get; set; } = default!;

        public string FileName
        {
            get
            {
                return _fileName ?? default!; 
            }
            set
            {
                _fileName= value;
                (int debutSprint, int finSprint) semainesSprint = InfoSprint.ExtractSemainesSprint(value);
                NumeroSemaineDebutDeSprint = semainesSprint.debutSprint;
                NumeroSemaineFinDeSprint = semainesSprint.finSprint;
            }
        }

        public string FullFileName { get; set; } = default!;

        public int NumeroSemaineDebutDeSprint { get; set; }
        public int NumeroSemaineFinDeSprint { get; set; }
       
        public TableauSuivi TabSuivi { get; set; } = new TableauSuivi();
    }

    public class TableauSuivi
    {
        public string SheetName { get; set; } = default!;
        public string TableName { get; set; } = default!;
        public ColumnTableauSuivi NumColumnTable { get; set; } = new ColumnTableauSuivi();
    }

    public class ColumnTableauSuivi
    {
        public int NoColumnApplication { get; set; } = default!;
        public int NoColumnDemande { get; set; } = default!;
        public int NoColumnHoursDevConsumed { get; set; }
        public int NoColumnHoursQualConsumed { get; set; }
    }
}