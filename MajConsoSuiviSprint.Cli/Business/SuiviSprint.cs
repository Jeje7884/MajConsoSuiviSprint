using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MajConsoSuiviSprint.Cli.Business.Interfaces;
using MajConsoSuiviSprint.Cli.Constants;
using MajConsoSuiviSprint.Cli.Model;
using MajConsoSuiviSprint.Cli.Utils;
using System;

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
                                                                                 );

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

        public void Test()
        {
           
             
                string lblLigneFin = "RESERVE";

                bool bTableau = true;
                int starrow = 5;
              
                string filePath = _configurationApp.SuiviSprintInfoConfig.FullFileName;
                int startRowIndex = 5; // Indice de la première ligne à partir de laquelle vous voulez lire


                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    string sheetNameToFind = "PI2023.08-1 Suivi Sprint-Init";
                   
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
    }
}