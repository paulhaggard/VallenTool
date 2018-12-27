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
        /// <returns>Returns the range created by this data, that was filled in on the worksheet.</returns>
        Excel.Range CreateData(Excel._Workbook workbook, int column_offset, int row_offset);

        /// <summary>
        /// Gets a string array that represents the data in the dataset
        /// </summary>
        /// <returns>Returns a 2-dimensional array of information stored in the dataset</returns>
        string[,] GetStringData();
    }
}
