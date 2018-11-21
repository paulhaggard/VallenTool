using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToolkit.DataFormatting;
using Microsoft.Office.Interop.Excel;

namespace ExcelToolkit
{
    public class ExcelWorkbookFormatter : IExcelData
    {
        public void CreateData(_Workbook workbook, int column_offset, int row_offset)
        {
            workbook.Worksheets.Delete();
            workbook.Worksheets.Add();
            workbook.Worksheets[1].Name = "Data";
        }
    }
}
