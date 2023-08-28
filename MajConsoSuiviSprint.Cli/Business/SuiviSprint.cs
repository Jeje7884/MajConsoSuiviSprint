using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
//using DocumentFormat.OpenXml.Wordprocessing;
using MajConsoSuiviSprint.Cli.Business.Interfaces;
using MajConsoSuiviSprint.Cli.Constants;
using MajConsoSuiviSprint.Cli.Model;
using MajConsoSuiviSprint.Cli.Utils;
using System.Text.RegularExpressions;

//using NPOI.SS.Formula.Functions;


namespace MajConsoSuiviSprint.Cli.Business
{
    internal class SuiviSprint
    {
        private readonly IConfigurationsApp _configurationApp;

        public SuiviSprint(IConfigurationsApp configuration)
        {
            _configurationApp = configuration;
        }

        public Dictionary<string, TempsConsommeDemandeModel> GetTempsConsommeesParDemandeSurUneSemaine(IList<ImportWebTTTExcelModel> saisiesActivitesInWebTTT)
        {
            var result = new Dictionary<string, TempsConsommeDemandeModel>();
            var saisiesActivitesSurLeSprint = saisiesActivitesInWebTTT
                                                                            .Where(data => data.NumeroDeSemaineDateActivite.Equals(_configurationApp.SuiviSprintInfoConfig.NumeroSemaineDebutDeSprint)
                                                                                                            ||
                                                                                                            data.NumeroDeSemaineDateActivite.Equals(_configurationApp.SuiviSprintInfoConfig.NumeroSemaineFinDeSprint)
                                                                                 )
                                                                            .OrderBy(data => data.NumeroDeDemande);

            if (saisiesActivitesSurLeSprint.Any())
            {
                foreach (var saisieActivite in saisiesActivitesSurLeSprint)
                {
                    string numeroFormatSuiviSprint = InfoSprint.ModifyDemandeWebTTTToSuiviSprint(saisieActivite.NumeroDeDemande, saisieActivite.Activite);
                    if (InfoSprint.IsActivityToManaged(saisieActivite.Activite))
                    {
                        if (result.ContainsKey(numeroFormatSuiviSprint))
                        {
                            if (saisieActivite.Activite.Equals(AppliConstant.LblActiviteQual))
                            {
                                result[numeroFormatSuiviSprint].HeureTotaleDeQualification += saisieActivite.HeureDeclaree;
                            }
                            else
                            {
                                result[numeroFormatSuiviSprint].HeureTotaleDeDeveloppement += saisieActivite.HeureDeclaree;
                            }
                        }
                        else
                        {
                            var conso = new TempsConsommeDemandeModel
                            {
                                Application = saisieActivite.Application,
                                HeureTotaleDeDeveloppement = saisieActivite.Activite.Equals(AppliConstant.LblActiviteDev) ? saisieActivite.HeureDeclaree : 0,
                                HeureTotaleDeQualification = saisieActivite.Activite.Equals(AppliConstant.LblActiviteQual) ? saisieActivite.HeureDeclaree : 0,
                                IsDemandeValide = InfoSprint.IsDemandeValid(saisieActivite, _configurationApp.WebTTTInfoConfig.ReglesSaisiesAutorisesParActivite)
                            };
                            result.Add(numeroFormatSuiviSprint, conso);
                        }
                    }
                }
            }

            var sortedList = result.ToList();
            sortedList.Sort(comparison: (pair1, pair2) => (pair1.Value.ToString() ?? "").CompareTo(pair2.Value));
            return result;
        }

        /// <summary>
        /// méthode ne fonctionne pas car la position du temps conso change
        /// car le nombre de colonne connu correspond au nombre de cellules renseignés par ligne
        /// </summary>
        /// <param name="dataSaisie"></param>
        public void UpdateFichierSuiviSprint(Dictionary<string, TempsConsommeDemandeModel> dataSaisie)
        {
            if (Divers.IsFileExist(_configurationApp.SuiviSprintInfoConfig.FullFileName))
            {


                string filePath = _configurationApp.SuiviSprintInfoConfig.FullFileName;



                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;

                    foreach (WorksheetPart wsp in workbookPart.WorksheetParts)
                    {
                        Worksheet ws = wsp.Worksheet;

                        WorksheetPart targetWorksheetPart = workbookPart.WorksheetParts.First();
                        // i want to do something like this

                        SheetData sheetData = ws.GetFirstChild<SheetData>();
                        int rowEC = 0;
                        foreach (Row row in sheetData.Elements<Row>())
                        {
                            if (rowEC > 4)
                            {
                                int column = 0;
                                foreach (Cell cell in row.Elements<Cell>())
                                {

                                    if (column == 31)
                                    {
                                        string cellValue = GetCellValue(cell, workbookPart);
                                        // Console.WriteLine($"N° demande {cellValue}\t");
                                    }

                                    column++;
                                }


                            }
                            rowEC++;
                        }
                    }

                }

            }
        }

        public void TestLectureToutOngletSuiviSprint()
        {
            string filePath = _configurationApp.SuiviSprintInfoConfig.FullFileName;


            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                //string sheetNameToFind = "PI2023.08-1 Suivi Sprint-Init";
                var sheets = workbookPart?.Workbook?.Sheets?.Cast<Sheet>().ToList();

                sheets?.ForEach(x => Console.WriteLine(
                      String.Format("RelationshipId:{0}\n SheetName:{1}\n SheetId:{2}"
                      , x?.Id?.Value, x?.Name?.Value, x?.SheetId?.Value)));

                foreach (WorksheetPart wsp in workbookPart?.WorksheetParts)
                {


                    Worksheet ws = wsp.Worksheet;


                    string partRelationshipId = workbookPart.GetIdOfPart(wsp);


                    SheetData sheetData = ws.GetFirstChild<SheetData>();


                    int rowEC = 0;
                    foreach (Row row in sheetData.Elements<Row>())
                    {
                        if (rowEC > 4)
                        {
                            int column = 0;
                            foreach (Cell cell in row.Elements<Cell>())
                            {

                                if (column == 31)
                                {
                                    string cellValue = GetCellValue(cell, workbookPart);
                                    Console.WriteLine($"N° demande {cellValue}\t");
                                }

                                column++;
                            }


                        }
                        rowEC++;
                    }
                }

            }
        }

        public void TestLectureTableauSuiviSprint()
        {



            string filePath = _configurationApp.SuiviSprintInfoConfig.FullFileName;


            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;

                IEnumerable<Sheet> sheets =
                workbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name.Value == "PI2023.08-1 Suivi Sprint");

                string relationshipId = sheets?.First().Id.Value; //rId2 ou rId1 

                WorksheetPart MyworksheetPart = (WorksheetPart)workbookPart.GetPartById(relationshipId);
                TableDefinitionPart tableDefinitionPart = MyworksheetPart.TableDefinitionParts.FirstOrDefault(t => t.Table.Name == "TableauSuiviSprint");
                if (tableDefinitionPart != null)
                {

                    Table table = tableDefinitionPart.Table;

                    string reference = table.Reference;
                    string[] parts = reference.Split(':');
                    string startCellReference = parts[0];
                    string endCellReference = parts[1];
                    // Get the worksheet data
                    Worksheet ws = MyworksheetPart.Worksheet;
                    SheetData sheetData = ws.GetFirstChild<SheetData>();
                    //SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                    int rowEC = 0;

                    //int columnNumDemande = 31;
                    int columnNumDemande = 33;
                    int columnNumDemandeReelle = columnNumDemande;
                    int nbColumnTotal = 55;
                    foreach (Row row in sheetData.Elements<Row>())
                    {

                        if (rowEC >= 4)
                        {
                            Cell firstCell = row.Elements<Cell>().ElementAtOrDefault(35);

                            if (firstCell != null)
                            {
                                string cellReference = firstCell.CellReference.Value;

                                if (string.Compare(cellReference, startCellReference) >= 0 && string.Compare(cellReference, endCellReference) <= 0)
                                {
                                    var nbCell = row.Elements<Cell>().Count();
                                    //var decalage = nbColumnTotal - nbCell;
                                    var decalage = GetNbColumnDecalage(nbColumnTotal, nbCell);
                                    //if (decalage >0)
                                    //{
                                    //    columnNumDemandeReelle = columnNumDemande - (decalage+1);
                                    //}
                                    //else if((decalage < 0))
                                    //{
                                    //    columnNumDemandeReelle = columnNumDemande + decalage;
                                    //}
                                    //else
                                    //{
                                    //    columnNumDemandeReelle = columnNumDemande-1;
                                    //}
                                    if (rowEC == 29)
                                    {
                                        Console.Write("test");
                                        //cellDemandeAffecte = row.Elements<Cell>().ElementAtOrDefault(28);
                                        //string cellValue = GetCellValue(cellDemande);
                                        //cellValueAff = GetCellValue(cellDemandeAffecte, workbookPart);
                                    }
                                    //int numAffect;
                                    if (decalage >= 0)
                                    {
                                        // numAffect = columnNumAffec + decalage;
                                        columnNumDemandeReelle = columnNumDemande - decalage;
                                    }
                                    else
                                    {
                                        //numAffect = columnNumAffec - decalage;
                                        columnNumDemandeReelle = columnNumDemande + decalage;
                                    }
                                    //cellDemandeAffecte = row.Elements<Cell>().ElementAtOrDefault(numAffect);
                                    //cellValueAff = GetCellValue(cellDemandeAffecte, workbookPart);

                                    //Console.WriteLine($"valeur affecté  {cellValueAff}\t nb cell : {nbCell} numLigne {rowEC}");
                                    //if (decalage >0)
                                    //{
                                    //    columnNumDemandeReelle = columnNumDemande - (decalage+1);
                                    //}
                                    //else if((decalage < 0))
                                    //{
                                    //    columnNumDemandeReelle = columnNumDemande + decalage;
                                    //}
                                    //else
                                    //{
                                    //    columnNumDemandeReelle = columnNumDemande-1;
                                    //}
                                    //if (rowEC ==29)
                                    //{
                                    //    Console.Write("test");
                                    //    cellDemandeAffecte = row.Elements<Cell>().ElementAtOrDefault(28);
                                    //    //string cellValue = GetCellValue(cellDemande);
                                    //    cellValueAff = GetCellValue(cellDemandeAffecte, workbookPart);
                                    //}


                                    //if (cellValueAff!="")

                                    //{
                                    //    columnNumDemandeReelle = columnNumDemandeReelle + 1;
                                    //}

                                    //if (rowEC == 4)
                                    //{
                                    //    columnNumDemandeReelle = columnNumDemande - 2;
                                    //}
                                    //else
                                    //{
                                    //    columnNumDemandeReelle = columnNumDemande;
                                    //}

                                    //if (row.Elements<Cell>().Count())
                                    //{

                                    //}
                                    var cellDemande = row.Elements<Cell>().ElementAtOrDefault(columnNumDemandeReelle);
                                    //string cellValue = GetCellValue(cellDemande);
                                    string cellValue = GetCellValue(cellDemande, workbookPart);
                                    Console.WriteLine($"valeur  {cellValue}\t nb cell : {nbCell} numLigne {rowEC}");


                                }
                                //foreach (Cell cell in row.Elements<Cell>())
                                //{


                                //    string cellValue = GetCellValue(cell, workbookPart);
                                //    Console.WriteLine($"valeur  {cellValue}\t column : {column}");
                                //    //if (column == 31)
                                //    //{
                                //    //    string cellValue = GetCellValue(cell, workbookPart);
                                //    //    Console.WriteLine($"N° demande {cellValue}\t");
                                //    //}

                                //    column++;
                                //}


                            }
                        }
                        rowEC++;
                    }
                }
            }
        }

        /// <summary>
        /// mise à jour de la colonne "AR" qd le num demande est trouvée
        /// 
        /// </summary>
        public void TestUpdateTableauSuiviSprintByRefColumn(bool couldUpdateColumnConso)
        {

            string filePath = _configurationApp.SuiviSprintInfoConfig.FullFileName;

            using SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, true);

            WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;

            //   IEnumerable<Sheet> sheets =
            //    workbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s? .Name?.Value == "PI2023.08-1 Suivi Sprint");
            IEnumerable<Sheet> sheets =
            workbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();

            string relationshipId = sheets.First().Id.Value; //récupération du 1er onglet


            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(relationshipId);
            TableDefinitionPart tableDefinitionPart = worksheetPart.TableDefinitionParts
                                                                                .FirstOrDefault(t => t.Table.Name == "TableauSuiviSprint");

            if (tableDefinitionPart != null)
            {

                Table table = tableDefinitionPart.Table;

                string reference = table.Reference;
                string[] parts = reference.Split(':');
                string startCellReference = parts[0];
                string endCellReference = parts[1];
                // Get the worksheet data
                Worksheet ws = worksheetPart.Worksheet;
                SheetData sheetData = ws.GetFirstChild<SheetData>();
                int rowEC = 5;
                int nbLigneMaxTableau= int.Parse(Regex.Match(endCellReference, @"\d+").Value);

                foreach (Row row in sheetData.Elements<Row>().Skip(4))
                {
                    if (rowEC > nbLigneMaxTableau)

                    {
                        break;
                    }

                    Cell firstCell = row.Elements<Cell>().ElementAtOrDefault(35);// colonne de rérérence du tableau

                    if (firstCell != null)
                    {
                        string cellReference = firstCell.CellReference.Value;
                        if (rowEC>=41)
                        {
                            Console.WriteLine("ic");
                        }

                        if (string.Compare(cellReference, startCellReference) >= 0 && string.Compare(cellReference, endCellReference) <= 0)
                        {
                            var nbCell = row.Elements<Cell>().Count();
                            string cellReferenceNumDemande = $"AH{rowEC}";

                            string? cellValue;

                            var celltest = row.Elements<Cell>().Where(s => s.CellReference == cellReferenceNumDemande);
                            if (celltest.Any())
                            {
                                cellValue = GetCellValue(celltest.First(), workbookPart);
                            }
                            else
                            {

                                Cell? cellNumDemande = worksheetPart.Worksheet.Descendants<Cell>()
                                    .FirstOrDefault(c => c?.CellReference?.Value == cellReferenceNumDemande);

                                if(cellNumDemande != null)
                                {
                                    cellValue = GetCellValue(cellNumDemande);
                                }
                                else
                                {
                                    cellValue = null;
                                }
                                
                            }                   

                            Console.WriteLine($"La valeur de la cellule demande {cellReferenceNumDemande} est : {cellValue}");
                            if (couldUpdateColumnConso & cellValue != null)
                            {
                              
                                var cellsheureConso = row.Elements<Cell>().Where(s => s.CellReference == $"AR{rowEC}");
                                Cell cellheureConso = cellsheureConso.First();
                                cellheureConso ??= InsertCellInWorksheet(ws, $"AR{rowEC}");
                                
                                if (cellValue=="#2")
                                {
                                    cellheureConso.CellValue = new CellValue("Erreur : 5");
                                }
                                else
                                {
                                    cellheureConso.CellValue = new CellValue("5");
                                }
                                cellheureConso.DataType = new EnumValue<CellValues>(CellValues.String);
                            }
   
                        }

                    }
                    rowEC++;
                }


            }
            spreadsheetDocument.Save();
        }



        private static Cell InsertCellInWorksheet(Worksheet worksheet, string cellReference)
        {
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            Cell cell = new() { CellReference = cellReference };
            sheetData.Append(cell);
            return cell;
        }

        /// <summary>
        /// Méthod qui peut marcher si on crée un onglet dédié, référençant le temps par demande
        /// </summary>
        public void TestUpdateTableauConsoInNewTab()
        {


            string filePath = _configurationApp.SuiviSprintInfoConfig.FullFileName;


            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, true))
            {
                WorkbookPart? workbookPart = spreadsheetDocument.WorkbookPart;

                if (null != workbookPart)
                {
                    IEnumerable<Sheet>? sheets =
                                                workbookPart?.Workbook?
                                                .GetFirstChild<Sheets>()?
                                                .Elements<Sheet>()?
                                                .Where(s => s?.Name?.Value == "Conso");

                    string? relationshipId = sheets?.First().Id?.Value; //rId2 ou rId1 

                    WorksheetPart? worksheetPart = (WorksheetPart)workbookPart.GetPartById(relationshipId ?? "rd1") ?? null;

                    TableDefinitionPart? tableDefinitionPart = worksheetPart?.TableDefinitionParts?.FirstOrDefault(t => t?.Table?.Name == "TableauConso");

                    if (tableDefinitionPart != null && worksheetPart != null)
                    {

                        Table table = tableDefinitionPart.Table;

                        ClearTableData(tableDefinitionPart, worksheetPart);

                        table.Reference = "A1:C2";
                        table.Save();

                        Dictionary<string, TempsConsommeDemandeModel> dataSaisie = new();

                        var saisie = new TempsConsommeDemandeModel()
                        {
                            HeureTotaleDeDeveloppement = 10,
                            HeureTotaleDeQualification = 20
                        };
                        dataSaisie.Add("12596", saisie);

                        saisie = new TempsConsommeDemandeModel()
                        {
                            HeureTotaleDeDeveloppement = 10,
                            HeureTotaleDeQualification = 2
                        };
                        dataSaisie.Add("12599", saisie);

                        saisie = new TempsConsommeDemandeModel()
                        {
                            HeureTotaleDeDeveloppement = 30,
                            HeureTotaleDeQualification = 50,
                            IsDemandeValide = false
                        };
                        dataSaisie.Add("9999", saisie);

                        foreach (var item in dataSaisie)
                        {
                            // SheetData sheetData = tableDefinitionPart.GetFirstChild<Table>().Elements<SheetData>().First();
                            SheetData? sheetData = worksheetPart?.Worksheet?.GetFirstChild<SheetData>();
                            Row newRow = new();
                            Cell cell1 = new(new CellValue(item.Key));
                            var valKey = item.Key;
                            var valValue = item.Value.HeureTotaleDeQualification;

                            Cell cell2 = new(new CellValue(item.Value.HeureTotaleDeDeveloppement));

                            Cell cell3 = new(new CellValue(item.Value.HeureTotaleDeQualification));

                            newRow.Append(cell1, cell2, cell3);
                            sheetData?.Append(newRow);

                        }

                        table.Reference = $"A1:C{dataSaisie.Count + 1}";
                        table.Save();

                        // Save the changes
                        workbookPart.Workbook.Save();
                    }
                }
            }
        }



        private int GetNbColumnDecalage(int nbColumnTotal, int nbColumnNonVide)
        {
            int nbColumnDecalage = 0;
            var decalage = nbColumnTotal - nbColumnNonVide;
            if (decalage > 0)
            {
                nbColumnDecalage = -(decalage + 1);
            }
            else if ((decalage < 0))
            {
                nbColumnDecalage = decalage;
            }
            else
            {
                nbColumnDecalage = -1;
            }


            return nbColumnDecalage;
        }



        private static void ClearTableData(TableDefinitionPart tableDefinitionPart, WorksheetPart ws)
        {
            Table table = tableDefinitionPart.Table;
            //   var relId = worksheetPart.GetIdOfPart(tableDefinitionPart);
            // WorksheetPart worksheetPart = (WorksheetPart)table.Parent;
            SheetData sheetData = ws.Worksheet.GetFirstChild<SheetData>();

            // Row lastRow = sheetData.Elements<Row>().LastOrDefault();
            // Get the worksheet data
            int nbRow = sheetData.Elements<Row>().Count();
            int numRow = 0;

            for (int i = 0; i < nbRow; i++)
            {

                var currentRow = sheetData.Elements<Row>().ElementAt(i);
                if (i > 0)
                {

                    currentRow.Remove();

                    // }

                }
                numRow++;


            }

            // Remove all rows in the table except for the header row
            //foreach (Row row in sheetData.Elements<Row>())
            //{
            //    if (numRow > 0)
            //    {
            //        Cell firstCell = row.Elements<Cell>().FirstOrDefault();
            //        if (firstCell != null)
            //        {
            //            string cellReference = firstCell.CellReference.Value;

            //            // Check if the cell is within the table range
            //            // if (string.Compare(cellReference, table.Reference.Value) > 0)
            //            // {
            //            //int nbColumns = row.Elements<Cell>().Count();
            //            //for (int i = 0; i <= nbColumns; i++)
            //            //{

            //            //}
            //            // Remove the row
            //            row.Remove();

            //            // }
            //        }
            //    }
            //    numRow++;


            //}

            //  tableDefinitionPart.Table.Save();//moi qui est ajouté

        }



        private static string GetCellValue(Cell cell)
        {
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return cell.CellValue.Text;
            }
            else
            {
                return cell.InnerText;
            }


        }

        private static string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                int sharedStringIndex = int.Parse(cell.CellValue.Text);
                SharedStringTablePart sharedStringTablePart = workbookPart.SharedStringTablePart;
                return sharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(sharedStringIndex).InnerText;
            }
            else
            {
                return cell.InnerText;
            }
        }


        private void AddRowsToTable(TableDefinitionPart tableDefinitionPart)
        {
            // Get the worksheet data
            //SheetData sheetData =tableDefinitionPart.GetFirstChild<Table>().Elements<SheetData>().First();
            SheetData sheetData2 = tableDefinitionPart.Table.Elements<SheetData>().First();

            // Create a new row
            Row newRow = new();

            // Add cells to the row with appropriate values
            Cell cell1 = new(new CellValue("Value1")); // Column 1
            Cell cell2 = new(new CellValue("Value2")); // Column 2

            newRow.Append(cell1, cell2);

            // Append the row to the sheet data
            sheetData2.Append(newRow);
        }
        private void AddRows2ToTable()
        {
            string filename = _configurationApp.SuiviSprintInfoConfig.FullFileName;
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filename, true))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == "Sheet1").First();
                WorksheetPart worksheetPart = workbookPart.GetPartById(sheet.Id) as WorksheetPart;
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                //write some new data
                for (int i = 1; i < 11; i++)
                {
                    //rowIndex takes into account the header (i.e. we're filling A2-A11)
                    uint rowIndex = (uint)i + 1;
                    Row row = new() { RowIndex = rowIndex };
                    Cell cell = new() { CellReference = $"A{rowIndex}", CellValue = new CellValue(i), DataType = CellValues.Number };
                    row.Append(cell);

                    sheetData.Append(row);
                }
                //find the table
                Table table = worksheetPart.TableDefinitionParts.FirstOrDefault(t => t.Table.Name == "MyTable")?.Table;
                //update the reference of the table (11 is 10 data rows and 1 header)
                table.Reference = "A1:A11";
            }
        }


        private static void CreateTable(WorksheetPart worksheetPart)
        {

            List<object> list = new() { new { Id = 001, FirstName = "John", LastName = "Doe", Department = "Marketing" }, new { Id = 002, FirstName = "Jane", LastName = "Doe", Department = "Accounting" } };
            // Create the table definition and properties
            TableDefinitionPart tableDefinitionPart = worksheetPart.AddNewPart<TableDefinitionPart>();
            Table table = new()
            {
                Id = 1,
                Name = "TableName",
                DisplayName = "Table Display Name",
                Reference = "A1:C10", // Adjust the range as needed
                TotalsRowCount = 0
            };
            table.AppendChild(new AutoFilter() { Reference = table.Reference });

            // Attach the table to the worksheet
            tableDefinitionPart.Table = table;

            // Save the changes
            tableDefinitionPart.Table.Save();
        }

        private static void DefineTable(WorksheetPart worksheetPart, int rowMin, int rowMax, int colMin, int colMax)
        {
            TableDefinitionPart tableDefinitionPart = worksheetPart.AddNewPart<TableDefinitionPart>("rId" + (worksheetPart.TableDefinitionParts.Count() + 1));
            int tableNo = worksheetPart.TableDefinitionParts.Count();

            string reference = ((char)(64 + colMin)).ToString() + rowMin + ":" + ((char)(64 + colMax)).ToString() + rowMax;

            Table table = new() { Id = (UInt32)tableNo, Name = "Table" + tableNo, DisplayName = "Table" + tableNo, Reference = reference, TotalsRowShown = false };
            AutoFilter autoFilter = new() { Reference = reference };

            TableColumns tableColumns = new() { Count = (UInt32)(colMax - colMin + 1) };
            for (int i = 0; i < (colMax - colMin + 1); i++)
            {
                tableColumns.Append(new TableColumn() { Id = (UInt32)(i + 1), Name = "Column" + i });
            }

            TableStyleInfo tableStyleInfo = new() { Name = "TableStyleLight1", ShowFirstColumn = false, ShowLastColumn = false, ShowRowStripes = true, ShowColumnStripes = false };

            table.Append(autoFilter);
            table.Append(tableColumns);
            table.Append(tableStyleInfo);

            tableDefinitionPart.Table = table;

            TableParts tableParts = new() { Count = 1 };
            TablePart tablePart = new() { Id = "rId" + tableNo };

            tableParts.Append(tablePart);

            worksheetPart.Worksheet.Append(tableParts);
        }
    }
}