using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace AxOpenXML_Part1
{
    class Program
    {
        static void Main(string[] args)
        {
            // Generate a random temp file name
            string tempFile = System.IO.Path.GetTempFileName();
            string tempFilePath = System.IO.Path.GetDirectoryName(tempFile);
            string tempFileName = System.IO.Path.GetFileNameWithoutExtension(tempFile);
            string filename = tempFilePath + tempFileName + 
                              ".xlsx";

            // Create the file
            CreateWorkbook(filename);

            // Open the file in excel
            System.Diagnostics.Process.Start(filename);
        }

        /// Creates the workbook
        public static SpreadsheetDocument CreateWorkbook(string fileName)
        {
            SpreadsheetDocument spreadSheet = null;
            WorkbookStylesPart workbookStylesPart;

            try
            {
                // Create the Excel workbook
                using (spreadSheet = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook, false))
                {
                    // Create the parts and the corresponding objects
                    // Workbook
                    var workbookPart = spreadSheet.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();

                    // WorkSheet
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    worksheetPart.Worksheet = new Worksheet(sheetData);
                    var sheets = spreadSheet.WorkbookPart.Workbook.AppendChild(new Sheets());
                    var sheet = new Sheet()
                    {
                        Id = spreadSheet.WorkbookPart
                            .GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "Sheet 1"
                    };
                    sheets.AppendChild(sheet);

                    // Stylesheet                    
                    workbookStylesPart = spreadSheet.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                    workbookStylesPart.Stylesheet = new Stylesheet();
                    workbookStylesPart.Stylesheet.Save();

                    // Write some values
                    WriteSomeValues(worksheetPart);

                    // Save the workbook
                    worksheetPart.Worksheet.Save();
                    spreadSheet.WorkbookPart.Workbook.Save();
                }

            }
            catch (System.Exception exception)
            {
                Console.WriteLine(exception.Message);
            }

            return spreadSheet;
        }

        private static void WriteSomeValues(WorksheetPart worksheetPart)
        {
            int numRows = 5;
            int numCols = 3;

            for (int row = 0; row < numRows; row++)
            {
                Row r = new Row();
                for (int col = 0; col < numCols; col++)
                {
                    Cell c = new Cell();
                    CellFormula f = new CellFormula();
                    f.CalculateCell = true;
                    f.Text = "RAND()";
                    c.Append(f);
                    CellValue v = new CellValue();
                    c.Append(v);
                    r.Append(c);
                }

                worksheetPart.Worksheet.GetFirstChild<SheetData>().Append(r);
            }
        }

    }
}
