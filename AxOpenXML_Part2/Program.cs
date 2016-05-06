using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using AxOpenXML_Part2_dll;

namespace AxOpenXML_Part2
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

            ExcelExport excelExportObj = new ExcelExport();
            excelExportObj.exportProjTable(filename);

            // Open the file in excel
            System.Diagnostics.Process.Start(filename);
        }
    }
}
