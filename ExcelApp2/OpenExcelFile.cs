using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelApp2
{
    public static class OpenExcelFile
    {
        public static void getExcelFile(out System.Array X,  out System.Array date)
        {
            var path = "";

            // Set path to excel file
            OpenFileDialog dialog = new OpenFileDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
               path = dialog.FileName;
            }

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@path);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlXRange = xlWorksheet.UsedRange.Columns[5];
            Excel.Range xlYRange = xlWorksheet.UsedRange.Columns[3];
            X = (System.Array)xlXRange.Cells.Value;
            date = (System.Array)xlYRange.Cells.Value;

          
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //close and release
            xlWorkbook.Close();

            //quit and release
            xlApp.Quit();


        }
    }
}
