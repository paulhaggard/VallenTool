using ExcelToolkit.DataFormatting;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToolkit
{
    public class ExcelBinPrinter : IExcelData
    {
        private ICollection<Dataset> Data { get; set; } = new List<Dataset>();
        private int rowDataLength { get => getRowDataLength(); }
        private int columnDataLength { get => getColumnDataLength(); }

        #region Getters

        private int getColumnDataLength()
        {
            int columnDataLength = 0;
            foreach (Dataset set in Data)
                if (set.Responses.Count > columnDataLength)
                    columnDataLength = set.Responses.Count;
            return columnDataLength;
        }

        private int getRowDataLength()
        {
            int rowDataLength = 0;
            foreach (Dataset set in Data)
                if (set.Frequencies.Count > rowDataLength)
                    rowDataLength = set.Frequencies.Count;
            return rowDataLength;
        }

        #endregion

        public ExcelBinPrinter(ICollection<Dataset> data, int numBins)
        {
            Data = Manip.GenerateFrequencyBins(data, numBins);
        }

        public virtual Excel.Range CreateData(Excel._Workbook workbook, int column_offset, int row_offset)
        {
            return CreateData(workbook, column_offset, row_offset, true);
        }

        public virtual Excel.Range CreateData(Excel._Workbook workbook, int column_offset, int row_offset, bool createChart = true)
        {
            Excel._Worksheet data = workbook.Worksheets["Frequency Bins"];

            // Gets the range from the current worksheet
            string columnLetter = ExcelPort.ColumnNumToColumnString(1);
            string nextColumnLetter = ExcelPort.ColumnNumToColumnString(columnDataLength + 1);
            Excel.Range range = data.Range[columnLetter + "1", nextColumnLetter + (Data.Count + 3 + rowDataLength)];


            // Writes the data to excel
            range.Value = GetStringData();

            if(createChart)
            {
                /* TODO
                 * Create method to extract XY coordinate pairs from the data.
                 */

                CreateChart(data, range, 4, 1);

            }

            return range;
        }

        /// <summary>
        /// Creates a chart on the given worksheet with the given data, at the given location with the given size
        /// </summary>
        /// <param name="sheet">The worksheet to put the chart onto</param>
        /// <param name="chartData">The data to put into the chart</param>
        /// <param name="column_offset">The column index to start the chart at</param>
        /// <param name="row_offset">The row index to start the chart at</param>
        /// <param name="height">The height of the chart in pixels</param>
        /// <param name="width">The width of the chart in pixels</param>
        public void CreateChart(Excel._Worksheet sheet, Excel.Range chartData, 
            int column_offset, int row_offset, double height = 300, double width = 300)
        {
            // Gets the range from the current worksheet
            string columnLetter = ExcelPort.ColumnNumToColumnString(column_offset);

            // Add chart.
            Excel.ChartObjects charts = sheet.ChartObjects();
            Excel.Range origin = sheet.Range[(columnLetter + row_offset)];
            Excel.ChartObject chartObject = charts.Add((double)origin.Top, (double)origin.Left, 300, 300);
            Excel.Chart chart = chartObject.Chart;

            // Sets the chart range
            chart.SetSourceData(chartData);

            // TODO: You probably need to change the units here
            // Set chart properties.
            chart.ChartType = Excel.XlChartType.xlLine;
            chart.ChartWizard(Source: chartData,
                Title: "Frequency vs. Responses",
                CategoryTitle: "Frequency (Hz)",
                ValueTitle: "Responses (dbm)");
        }

        public virtual string[,] GetStringData()
        {
            string[,] dt = new string[(Data.Count + rowDataLength + 3), columnDataLength + 1];

            dt[0, 0] = "Frequency Bin";
            dt[0, 1] = "Response Average";
            dt[0, 2] = "X1 Response Data";

            for (int i = 3; i <= columnDataLength; i++)
                dt[0, i] = "X" + (i - 1);

            for (int i = 2; i <= Data.Count; i++)
            {
                dt[i, 0] = "Bin " + (i - 1);

                for (int r = 0; r < Data.ElementAt(i - 2).Responses.Count; r++)
                    dt[i, r + 1] = Data.ElementAt(i - 2).Responses[r].ToString();

                dt[Data.Count + 2, i - 1] = "Bin " + (i - 1);

                for (int f = 0; f < Data.ElementAt(i - 2).Frequencies.Count; f++)
                    dt[f + Data.Count + 3, i - 2] = Data.ElementAt(i - 2).Frequencies[f].ToString();
            }

            return dt;
        }
    }
}
