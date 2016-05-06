using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MIL = Microsoft.Dynamics.AX.ManagedInterop;
using Microsoft.Dynamics.AX.Framework.Linq.Data;
using Microsoft.Dynamics.AX.Framework.Linq.Data.Common;
using Microsoft.Dynamics.AX.Framework.Linq.Data.ManagedInteropLayer;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Collections;

namespace AxOpenXML_Part2_dll
{
    public class ExcelExport
    {
        uint rowPos = 0;
        WorksheetPart worksheetpart;
        SharedStringTablePart shareStringPart;
        Columns columns;    

        public ExcelExport()
        {
            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(helper.CurrentDomain_AssemblyResolve);
            rowPos = 1;
        }

        public void exportProjTable(string filename)
        {            
            // Create a Session object and login if necessary
            if (MIL.Session.Current == null)
            {
                MIL.Session axSession = new MIL.Session();
                axSession.Logon(null, null, null, null);
            }

            string fullPath = filename;

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(fullPath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {

                // Add a new workbook part
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // WorkSheet
                worksheetpart = workbookpart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                worksheetpart.Worksheet = new Worksheet(sheetData);
                var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
                var sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart
                        .GetIdOfPart(worksheetpart),
                    SheetId = 1,
                    Name = "Sheet 1"
                };
                sheets.AppendChild(sheet);

                // Stylesheet                    
                WorkbookStylesPart workbookStylesPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                workbookStylesPart.Stylesheet = new Stylesheet();
                workbookStylesPart.Stylesheet.Save();

                // Get the SharedStringTablePart. If it does not exist, create a new one.                
                if (spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = spreadsheetDocument.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }

                // Insert Header
                InsertHeader();

                // Insert Data
                InsertTableData();

                // Set column width
                CreateColumns(worksheetpart);

                worksheetpart.Worksheet.Save();
                workbookpart.Workbook.Save();
            }
        }        

        private void InsertHeader()
        {
            InsertData(helper.GetLabel("@SYS11779"), "A");
            InsertData(helper.GetLabel("@SYS27874"), "B");
            InsertData(helper.GetLabel("@SYS11904"), "C");
            helper.MergeAdjCells(worksheetpart, new List<string>() { "C" + rowPos.ToString(), "D" + rowPos.ToString() });
            rowPos++;
        }

        private void InsertTableData()
        {
            QueryProvider provider = new AXQueryProvider(null);
            var custTableQuery = new QueryCollection<CustTable>(provider);
            var dirPartyQuery = new QueryCollection<DirPartyTable>(provider);
            var customerList = from ct in custTableQuery
                               join dp in dirPartyQuery
                               on ct.Party equals dp.RecId
                               orderby ct.AccountNum ascending
                               select new { ct.AccountNum, ct.CustGroup, dp.Name };

            foreach (var item in customerList)
            {
                InsertData(item.AccountNum, "A");
                InsertData(item.Name, "B");
                InsertData(item.CustGroup, "C");
                helper.MergeAdjCells(worksheetpart, new List<string>() { "C" + rowPos.ToString(), "D" + rowPos.ToString() });
                rowPos++;
            }
        }

        private void CreateColumns(WorksheetPart worksheetpart)
        {
            columns = new Columns();
            Column c = helper.CreateNewColumn(1U, 1U, 15);
            columns.Append(c);
            c = helper.CreateNewColumn(2U, 2U, 30);
            columns.Append(c);
            worksheetpart.Worksheet.InsertBefore(columns, worksheetpart.Worksheet.Elements<SheetData>().First());
        }

        private void InsertData(string value, string colPos, uint borderIndex = 0)
        {
            int index = helper.InsertSharedStringItem(value, shareStringPart);
            Cell newCell = helper.GetCellInWorksheet(colPos, rowPos, worksheetpart, true, false);
            newCell.CellValue = new CellValue(index.ToString());
            newCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            newCell.StyleIndex = borderIndex;
        }
    }
}
