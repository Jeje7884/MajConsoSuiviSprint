using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
//using DocumentFormat.OpenXml.Wordprocessing;
using MajConsoSuiviSprint.Cli.Business.Interfaces;
using MajConsoSuiviSprint.Cli.Constants;
using MajConsoSuiviSprint.Cli.Model;
using MajConsoSuiviSprint.Cli.Utils;
using DocumentFormat.OpenXml;
using System.Collections.Generic;


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
            
            var sort2= result.OrderBy(tps => tps.Key).ToList(); 
           // var sortedList = result.ToList();
            //sortedList.Sort((pair1, pair2) => pair1.Value.NumeroDeDemande(pair2.Value));
            return result;
        }

        public void UpdateFichierSuiviSprint(Dictionary<string, TempsConsommeDemandeModel> dataSaisie)
        {
            if (Divers.IsFileExist(_configurationApp.SuiviSprintInfoConfig.FullFileName))
            {
                //using ClosedXmlHelper closedXmlHelper = new(_configurationApp.SuiviSprintInfoConfig.FullFileName);
                //var sheet = closedXmlHelper.GetWorksheetByName(_configurationApp.SuiviSprintInfoConfig.TabSuivi.SheetName);

                //sheet ??= closedXmlHelper.GetWorksheetById(1);

                //var table = ClosedXmlHelper.GetTableByName(sheet, _configurationApp.SuiviSprintInfoConfig.TabSuivi.TableName);

                //foreach (var row in table.DataRange.Rows())
                //{
                //    int idUs = _configurationApp.SuiviSprintInfoConfig.TabSuivi.NumColumnTable.NoColumnDemande;
                //    string numDemande = ClosedXmlHelper.GetValueInCellFromTable(row, idUs);
                //    Console.WriteLine($"value us {numDemande} ");
                //    if (dataSaisie.ContainsKey(numDemande))
                //    {
                //        Console.WriteLine($"La valeur saisie est  {dataSaisie[numDemande].HeureTotaleDeDeveloppement}");
                //    }

                //}

                //using ExceLNPOIHelper exceLNPOIHelper = new(_configurationApp.SuiviSprintInfoConfig.FullFileName);

                // var sheet = exceLNPOIHelper.GetSheetByID(0); var rowHeaders = sheet.GetRow(3);

                //var numpUS = ExceLNPOIHelper.GetCellValue(rowHeaders, 30);
                string lblLigneFin = "RESERVE";

                bool bTableau = true;
                int starrow = 5;
                //if (ExceLNPOIHelper.GetCellValue(rowHeaders, 40).Equals("Heure tot DEV\n Consommé"))
                //{
                //    while (row <= sheet.LastRowNum && bTableau)
                //    {
                //        IRow dataRow = sheet.GetRow(row);
                //        string numpAppli = ExceLNPOIHelper.GetCellValue(dataRow, 30);
                //        int idColonneUS = 31;// _configurationApp.SuiviSprintInfoConfig.TabSuivi.NumColumnTable.NoColumnDemande;
                //        string numDemande = ExceLNPOIHelper.GetCellValue(dataRow, idColonneUS);
                //        string conso = ExceLNPOIHelper.GetCellValue(dataRow, 41);

                // for (int cellIndex = 0; cellIndex < dataRow.LastCellNum; cellIndex++) { ICell
                // cell = dataRow.GetCell(cellIndex); // Accédez à la cellule spécifiée

                // if (cell != null) { string cellValue = cell.ToString();
                // Console.WriteLine($"Valeur de la cellule [{dataRow},{cellIndex}] : {cellValue}");
                // } } Console.WriteLine($"value us {numDemande} nbe colonne {dataRow.Count()} -
                // valeur : {conso}" );

                // if (row == 10) { Console.WriteLine("te"); } if
                // (dataSaisie.ContainsKey(numDemande)) { Console.WriteLine($"La valeur saisie est
                // {dataSaisie[numDemande].HeureTotaleDeDeveloppement}"); } if
                // (ExceLNPOIHelper.GetCellValue(dataRow, 13).Contains(lblLigneFin)) {
                // bTableau=false; } row++; }

                //}

                string filePath = _configurationApp.SuiviSprintInfoConfig.FullFileName;
                int startRowIndex = 5; // Indice de la première ligne à partir de laquelle vous voulez lire

                //using (SpreadsheetDocument mySpreadsheet = SpreadsheetDocument.Open(filePath, false))

                //{
                //    var sheets = mySpreadsheet.WorkbookPart.Workbook.Sheets;

                // For each sheet, display the sheet information.

                // foreach (var sheet in sheets)

                // { var arttib = sheet.GetAttributes();

                // foreach (var attr in sheet.GetAttributes())

                // { if (attr.Value == "rId1") { break; } }

                // var rowList = sheet.Elements<Row>(); foreach (Row row in sheet.Elements<Row>())

                // { foreach (Cell cell in row.Elements<Cell>()) { string cellValue =
                // GetCellValue(cell); Console.Write($"{cellValue}\t"); }

                // Console.WriteLine(); // Nouvelle ligne pour la prochaine rangée }

                //    }
                //}

                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    string sheetNameToFind = "PI2023.08-1 Suivi Sprint-Init";
                    //    var sheetName = workbookPart.Workbook.Descendants<Sheet>().ElementAt(0).Name;
                    //WorksheetPart targetWorksheetPart = workbookPart.WorksheetParts.FirstOrDefault(
                    //    wsp => wsp.Worksheet.LocalName == sheetNameToFind);

                    foreach (WorksheetPart wsp in workbookPart.WorksheetParts)
                    {
                        Worksheet ws = wsp.Worksheet;
                        //WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                        WorksheetPart targetWorksheetPart = workbookPart.WorksheetParts.FirstOrDefault(
                            wsp => wsp.Worksheet.GetAttributes().Any(x => x.Value == sheetNameToFind));
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
                                        Console.WriteLine($"N° demande {cellValue}\t");
                                    }

                                    column++;
                                }


                            }
                            rowEC++;
                        }
                    }

                    // WorkbookPart workbookPart = doc.WorkbookPart; SharedStringTablePart sstpart =
                    // workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    // SharedStringTable sst = sstpart.SharedStringTable;

                    // WorksheetPart worksheetPart = workbookPart.WorksheetParts.First(); Worksheet
                    // sheet = worksheetPart.Worksheet; SheetData sheetData = sheet.GetFirstChild<SheetData>();

                    // var cells = sheetData.Descendants<Cell>(); var rows = sheetData.Descendants<Row>();

                    // Console.WriteLine("Row count = {0}", rows.LongCount());
                    // Console.WriteLine("Cell count = {0}", cells.LongCount());

                    // // One way: go through each cell in the sheet foreach (Cell cell in cells) {
                    // if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString)) {
                    // int ssid = int.Parse(cell.CellValue.Text); string str =
                    // sst.ChildElements[ssid].InnerText; Console.WriteLine("Shared string {0}:
                    // {1}", ssid, str); } else if (cell.CellValue != null) {
                    // Console.WriteLine("Cell contents: {0}", cell.CellValue.Text); } }

                    //    // Or... via each row
                    //    foreach (Row row in rows)
                    //    {
                    //        foreach (Cell c in row.Elements<Cell>())
                    //        {
                    //            if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
                    //            {
                    //                int ssid = int.Parse(c.CellValue.Text);
                    //                string str = sst.ChildElements[ssid].InnerText;
                    //                Console.WriteLine("Shared string {0}: {1}", ssid, str);
                    //            }
                    //            else if (c.CellValue != null)
                    //            {
                    //                Console.WriteLine("Cell contents: {0}", c.CellValue.Text);
                    //            }
                    //        }
                    //    }
                    //}
                }

                //using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(_configurationApp.SuiviSprintInfoConfig.FullFileName, false))
                //{
                //    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                //    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                //    int startRowIndex = 5;
                //    Worksheet worksheet = worksheetPart.Worksheet;
                //    SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                //    int rowIndex = 1; // Indice pour le suivi des lignes (commence à 1)
                //    var val = sheetData.Elements<Row>();

                // foreach (Row rowdata in sheetData.Elements<Row>()) { if (rowIndex >=
                // startRowIndex) { Cell cell = rowdata.Elements<Cell>().ElementAtOrDefault(31);
                // //foreach (Cell cell in rowdata.Elements<Cell>()) //{ // // Ici, vous pouvez
                // traiter les valeurs de cellule comme nécessaire // string cellValue =
                // GetCellValue(cell); // Console.Write($"{cellValue}\t"); //}

                // // Ici, vous pouvez traiter les valeurs de cellule comme nécessaire string
                // cellValue = GetCellValue(cell); Console.Write($"{cellValue}\t");
                // Console.WriteLine(); // Nouvelle ligne pour la prochaine rangée } rowIndex++; }

                //}
            }
        }

        public void TestLectureToutOngletSuiviSprint()
        {
           
             
                string lblLigneFin = "RESERVE";

                bool bTableau = true;
                int starrow = 5;
              
                string filePath = _configurationApp.SuiviSprintInfoConfig.FullFileName;
                int startRowIndex = 5; // Indice de la première ligne à partir de laquelle vous voulez lire


                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                //string sheetNameToFind = "PI2023.08-1 Suivi Sprint-Init";
                var sheets = workbookPart.Workbook.Sheets.Cast<Sheet>().ToList();

                sheets.ForEach(x => Console.WriteLine(
                      String.Format("RelationshipId:{0}\n SheetName:{1}\n SheetId:{2}"
                      , x.Id.Value, x.Name.Value, x.SheetId.Value)));

                foreach (WorksheetPart wsp in workbookPart.WorksheetParts)
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
                spreadsheetDocument.Save();
            }
        }

        public void TestLectureTableauSuiviSprint()
        {



            string filePath = _configurationApp.SuiviSprintInfoConfig.FullFileName;
            int startRowIndex = 5; // Indice de la première ligne à partir de laquelle vous voulez lire


            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;

                IEnumerable<Sheet> sheets =
                workbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name.Value == "PI2023.08-1 Suivi Sprint-Init");

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

                    foreach (Row row in sheetData.Elements<Row>())
                    {
                        Cell firstCell = row.Elements<Cell>().FirstOrDefault();

                        if (firstCell != null)
                        {
                            string cellReference = firstCell.CellReference.Value;

                            if (string.Compare(cellReference, startCellReference) >= 0 && string.Compare(cellReference, endCellReference) <= 0)
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

;
                            }
                        }
                    }
                }
            }
        }

        public void TestLectureTableauConso2()
        {



            string filePath = _configurationApp.SuiviSprintInfoConfig.FullFileName;
            int startRowIndex = 5; // Indice de la première ligne à partir de laquelle vous voulez lire

          
            string tableNameToClear = "TableName";

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, true))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.FirstOrDefault();

                if (worksheetPart != null)
                {
                    TableDefinitionPart tableDefinitionPart = worksheetPart.TableDefinitionParts.FirstOrDefault(t => t.Table.Name == tableNameToClear);

                    if (tableDefinitionPart != null)
                    {
                        // Clear existing rows in the table
                        ClearTableData(tableDefinitionPart, worksheetPart);
                    }
                }

                // Save the changes
                workbookPart.Workbook.Save();
            }

        }

        private static void ClearTableData(TableDefinitionPart tableDefinitionPart, WorksheetPart ws)
        {
            Table table = tableDefinitionPart.Table;
        //   var relId = worksheetPart.GetIdOfPart(tableDefinitionPart);
          // WorksheetPart worksheetPart = (WorksheetPart)table.Parent;
            SheetData sheetData = ws.Worksheet.GetFirstChild<SheetData>();

            Row lastRow = sheetData.Elements<Row>().LastOrDefault();
            // Get the worksheet data
            

            // Remove all rows in the table except for the header row
            foreach (Row row in sheetData.Elements<Row>())
            {
                Cell firstCell = row.Elements<Cell>().FirstOrDefault();

                if (firstCell != null)
                {
                    string cellReference = firstCell.CellReference.Value;

                    // Check if the cell is within the table range
                    if (string.Compare(cellReference, table.Reference.Value) > 0)
                    {
                        // Remove the row
                        row.Remove();
                    }
                }
            }
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


        public void TestAvecTableauConso()
        {
            string filePath = _configurationApp.SuiviSprintInfoConfig.FullFileName;
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, true))
        {
            WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
    WorksheetPart worksheetPart = workbookPart.WorksheetParts.FirstOrDefault();

            if (worksheetPart != null)
            {
                TableDefinitionPart tableDefinitionPart = worksheetPart.TableDefinitionParts.FirstOrDefault(t => t.Table.Name == "TableauConso");

                if (tableDefinitionPart != null)
                {
                    // Add rows to the table
                    AddRowsToTable(tableDefinitionPart);
}
            }

            // Save the changes
            workbookPart.Workbook.Save();
        }
        }
        private static void AddRowsToTable(TableDefinitionPart tableDefinitionPart)
        {
            // Get the worksheet data
            //SheetData sheetData =tableDefinitionPart.GetFirstChild<Table>().Elements<SheetData>().First();
            SheetData sheetData2 = tableDefinitionPart.Table.Elements<SheetData>().First();

            // Create a new row
            Row newRow = new Row();

            // Add cells to the row with appropriate values
            Cell cell1 = new Cell(new CellValue("Value1")); // Column 1
            Cell cell2 = new Cell(new CellValue("Value2")); // Column 2

            newRow.Append(cell1, cell2);

            // Append the row to the sheet data
            sheetData2.Append(newRow);
        }

        private static void CreateTable(WorksheetPart worksheetPart)
        {
            // Create the table definition and properties
            TableDefinitionPart tableDefinitionPart = worksheetPart.AddNewPart<TableDefinitionPart>();
            Table table = new Table()
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
    }
}