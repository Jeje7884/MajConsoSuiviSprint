using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Data;

namespace MajConsoSuiviSprint.Cli.Utils
{
    internal class testTable
    {
        public static DataTable GenericExcelTable(FileInfo fileName)
        {
            DataTable dataTable = new DataTable();
            try
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fileName.FullName, false))
                {
                    Workbook wkb = doc.WorkbookPart.Workbook;
                    Sheet wks = wkb.Descendants<Sheet>().FirstOrDefault();
                    SharedStringTable sst = wkb.WorkbookPart.SharedStringTablePart.SharedStringTable;
                    List<SharedStringItem> allSSI = sst.Descendants<SharedStringItem>().ToList();
                    WorksheetPart wksp = (WorksheetPart)doc.WorkbookPart.GetPartById(wks.Id);

                    foreach (TableDefinitionPart tdp in wksp.TableDefinitionParts)
                    {
                        QueryTablePart qtp = tdp.QueryTableParts.FirstOrDefault();
                        Table excelTable = tdp.Table;
                        int colcounter = 0;
                        foreach (TableColumn col in excelTable.TableColumns)
                        {
                            DataColumn dcol = dataTable.Columns.Add(col.Name);
                            dcol.SetOrdinal(colcounter);
                            colcounter++;
                        }

                        SheetData data = wksp.Worksheet.Elements<SheetData>().First();

                        foreach (Row row in data)
                        {
                            if (isInTable(row.Descendants<Cell>().FirstOrDefault(), excelTable.Reference, true))
                            {
                                int cellcount = 0;
                                DataRow dataRow = dataTable.NewRow();
                                foreach (Cell cell in row.Elements<Cell>())
                                {
                                    if (cell.DataType != null && cell.DataType.InnerText == "s")
                                    {
                                        dataRow[cellcount] = allSSI[int.Parse(cell.CellValue.InnerText)].InnerText;
                                    }
                                    else
                                    {
                                        dataRow[cellcount] = cell.CellValue.Text;
                                    }
                                    cellcount++;
                                }
                                dataTable.Rows.Add(dataRow);
                            }
                        }
                    }
                }
                //do whatever you want with the DataTable
                return dataTable;
            }
            catch (Exception ex)
            {
                //handle an error
                return dataTable;
            }
        }

        private static Tuple<int, int> returnCellReference(string cellRef)
        {
            int startIndex = cellRef.IndexOfAny("0123456789".ToCharArray());
            string column = cellRef.Substring(0, startIndex);
            int row = int.Parse(cellRef.Substring(startIndex));
            return new Tuple<int, int>(TextToNumber(column), row);
        }

        private static int TextToNumber(string text)
        {
            return text
                .Select(c => c - 'A' + 1)
                .Aggregate((sum, next) => sum * 26 + next);
        }

        private static bool isInTable(Cell testCell, string tableRef, bool headerRow)
        {
            Console.WriteLine("zzzz : " + tableRef);
            Tuple<int, int> cellRef = returnCellReference(testCell.CellReference.ToString());
            if (tableRef.Contains(":"))
            {
                int header = 0;
                if (headerRow)
                {
                    header = 1;
                }
                string[] tableExtremes = tableRef.Split(':');
                Tuple<int, int> startCell = returnCellReference(tableExtremes[0]);
                Tuple<int, int> endCell = returnCellReference(tableExtremes[1]);
                if (cellRef.Item1 >= startCell.Item1
                    && cellRef.Item1 <= endCell.Item1
                    && cellRef.Item2 >= startCell.Item2 + header
                    && cellRef.Item2 <= endCell.Item2) { return true; }
                else { return false; }
            }
            else if (cellRef.Equals(returnCellReference(tableRef)))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}