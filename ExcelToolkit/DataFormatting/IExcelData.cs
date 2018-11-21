using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToolkit.DataFormatting
{
    /// <summary>
    /// An interface for formatting data layouts of the datafiles passed into excel
    /// </summary>
    public interface IExcelData
    {
        /// <summary>
        /// Causes the implementing class to create data in the desired workbook in its desired fashion
        /// </summary>
        /// <param name="workbook">workbook to write to</param>
        /// <param name="column_offset">column offset (1='a'...)</param>
        /// <param name="row_offset">row offset</param>
        void CreateData(Excel._Workbook workbook, int column_offset, int row_offset);
    }
}
