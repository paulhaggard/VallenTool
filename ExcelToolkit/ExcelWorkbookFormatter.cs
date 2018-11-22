using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToolkit.DataFormatting;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToolkit
{
    public class ExcelWorkbookFormatter : IExcelData
    {
        public void CreateData(Excel._Workbook workbook, int column_offset, int row_offset)
        {
            // Deletes all but one worksheet from the workbook
            while(workbook.Worksheets.Count > 1)
            {
                workbook.Worksheets[2].Delete();
            }
            workbook.Worksheets[1].Name = "Data";
            workbook.Worksheets[1].Cells.Clear();

            //FORMATTING
            //Format headers as bold, vertical alignment = center.
            Excel._Worksheet sheet = workbook.Worksheets[1];

            // ERROR
            string columnLetter = ExcelPort.ColumnNumToColumnString(column_offset);

            sheet.get_Range("A1", columnLetter + "1").Font.Bold = true;
            sheet.get_Range("A1", columnLetter + "1").VerticalAlignment =
                Excel.XlVAlign.xlVAlignCenter;
            sheet.get_Range("A1", columnLetter + "1").ColumnWidth = 18;
        }

        public string[,] GetStringData()
        {
            throw new NotImplementedException();
        }
    }
}
