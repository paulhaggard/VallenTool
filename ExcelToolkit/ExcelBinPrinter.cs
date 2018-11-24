﻿using ExcelToolkit.DataFormatting;
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

        public void CreateData(Excel._Workbook workbook, int column_offset, int row_offset)
        {
            Excel._Worksheet data = workbook.Worksheets["Frequency Bins"];

            // Gets the range from the current worksheet
            string columnLetter = ExcelPort.ColumnNumToColumnString(1);
            string nextColumnLetter = ExcelPort.ColumnNumToColumnString(columnDataLength + 1);
            Excel.Range range = data.Range[columnLetter + "1", nextColumnLetter + (Data.Count + 3 + rowDataLength)];


            // Writes the data to excel
            range.Value = GetStringData();
        }

        public string[,] GetStringData()
        {
            string[,] dt = new string[(Data.Count + rowDataLength + 3), columnDataLength + 1];

            dt[0, 0] = "Frequency Bin";
            dt[0, 1] = "X1 Response Data";

            for (int i = 2; i <= columnDataLength; i++)
                dt[0, i] = "X" + i;

            for (int i = 1; i <= Data.Count; i++)
            {
                dt[i, 0] = "Bin " + i;

                for (int r = 0; r < Data.ElementAt(i - 1).Responses.Count; r++)
                    dt[i, r + 1] = Data.ElementAt(i - 1).Responses[r].ToString();

                dt[Data.Count + 2, i - 1] = "Bin " + i;

                for (int f = 0; f < Data.ElementAt(i - 1).Frequencies.Count; f++)
                    dt[f + Data.Count + 3, i - 1] = Data.ElementAt(i - 1).Frequencies[f].ToString();
            }

            return dt;
        }
    }
}
