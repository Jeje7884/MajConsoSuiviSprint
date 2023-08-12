namespace MajConsoSuiviSprint.Cli.Model
{
    public class SuiviSprintModel
    {
        public string FolderName { get; set; } = default!;
        public string FileName { get; set; } = default!;
        public string FullFileName { get; set; } = default!;

        public TableauSuivi TabSuivi { get; set; } = default!;
    }

    public class TableauSuivi
    {
        public string SheetName { get; set; } = default!;
        public string TableName { get; set; } = default!;
        public ColumnTableauSuivi NumColumnTable { get; set; } = default!;
    }

    public class ColumnTableauSuivi
    {
        public int NoColumnApplication { get; set; } = default!;
        public int NoColumnDemande { get; set; } = default!;
        public int NoColumnHoursDevConsumed { get; set; }
        public int NoColumnHoursQualConsumed { get; set; }
    }
}
