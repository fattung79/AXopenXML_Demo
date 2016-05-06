using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.Reflection;

namespace AxOpenXML_Part2_dll
{
    internal class helper
    {
        public helper() { }

        // Get Labels
        public static string GetLabel(string labelId, string langId = "en-us")
        {
            return SysLabel.labelId2String(labelId, langId);
        }

        // Translate Indexes into Column names (1 = A, 2 = B...27 = AA...703 = AAA...)
        public static string GetColNameFromIndex(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        // Translate Column names into Indexes (A = 1, B = 2...AA = 27...AAA = 703...)
        public static int GetColNumberFromName(string columnName)
        {
            char[] characters = columnName.ToUpperInvariant().ToCharArray();
            int sum = 0;
            for (int i = 0; i < characters.Length; i++)
            {
                sum *= 26;
                sum += (characters[i] - 'A' + 1);
            }
            return sum;
        }

        // Given a cell name, parses the specified cell to get the column name.
        public static string GetColumnName(string cellName)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);

            return match.Value;
        }

        // Given a cell name, parses the specified cell to get the row index.
        public static uint GetRowIndex(string cellName)
        {
            // Create a regular expression to match the row index portion the cell name.
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellName);

            return uint.Parse(match.Value);
        }

        // Given a column name, a row index, and a WorksheetPart,Get cell from the worksheet. 
        // If the cell doesn't exists, and createIfEmpty is true, create the cell and returns it.
        public static Cell GetCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart, Boolean createIfEmpty = true, Boolean saveWorksheet = false)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else if (createIfEmpty)
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    //if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    if (GetColNumberFromName(cell.CellReference.Value) > GetColNumberFromName(cellReference))
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                if (saveWorksheet)
                    worksheet.Save();
                return newCell;
            }
            else
            {
                return null;
            }
        }

        // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
        // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
        public static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }

        public static System.Reflection.Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            switch (args.Name)
            {                
                case "DocumentFormat.OpenXml, Version=2.5.5631.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35":
                    // Try to load ver 2.0 if 2.5 is not found.
                    return Assembly.Load("DocumentFormat.OpenXml, Version=2.0.5022.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35");
                case "DocumentFormat.OpenXml, Version=2.0.5022.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35":                    
                    // OpenXML2.0 and OpenXML2.5 both unavailable, throw error
                default:
                    throw new Exception("Can't find suitable OpenXML libraries for loading.");
            }
        }

        public static void MergeAdjCells(WorksheetPart worksheetPart, List<String> cells, Boolean saveWorksheet = false)
        {
            Worksheet worksheet = worksheetPart.Worksheet;

            foreach (String cName in cells)
            {
                string columnName = GetColumnName(cName);
                uint rowIndex = GetRowIndex(cName);
                GetCellInWorksheet(columnName, rowIndex, worksheetPart);
            }

            MergeCells mergeCells;
            if (worksheetPart.Worksheet.Elements<MergeCells>().Count() > 0)
            {
                mergeCells = worksheet.Elements<MergeCells>().First();
            }
            else
            {
                mergeCells = new MergeCells();

                // Insert a MergeCells object into the specified position.
                if (worksheet.Elements<CustomSheetView>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<CustomSheetView>().First());
                }
                else if (worksheet.Elements<DataConsolidate>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<DataConsolidate>().First());
                }
                else if (worksheet.Elements<SortState>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SortState>().First());
                }
                else if (worksheet.Elements<AutoFilter>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<AutoFilter>().First());
                }
                else if (worksheet.Elements<Scenarios>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<Scenarios>().First());
                }
                else if (worksheet.Elements<ProtectedRanges>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<ProtectedRanges>().First());
                }
                else if (worksheet.Elements<SheetProtection>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetProtection>().First());
                }
                else if (worksheet.Elements<SheetCalculationProperties>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetCalculationProperties>().First());
                }
                else
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
                }
            }

            MergeCell mergeCell = new MergeCell() { Reference = new StringValue(cells.First() + ":" + cells.Last()) };
            mergeCells.Append(mergeCell);

            if (saveWorksheet)
                worksheet.Save();
        }

        // Helper method to create Column objects
        public static Column CreateNewColumn(UInt32Value min, UInt32Value max, int width)
        {
            Column c = new Column()
            {
                Min = min,
                Max = max,
                Width = width,
                CustomWidth = true
            };

            return c;
        }

    }
}
