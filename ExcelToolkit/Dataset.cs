using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelToolkit.DataFormatting;

namespace ExcelToolkit
{
    /// <summary>
    /// A class used to contain a data file of frequencies and response data.
    /// </summary>
    public class Dataset : IExcelData, IDataManData<double>
    {
        #region Properties

        /// <summary>
        /// The filename, or just the name, used to load in this dataset and to identify this dataset among others
        /// </summary>
        public string Name { get; protected set; } = "";

        /// <summary>
        /// List of Frequencies for the test
        /// </summary>
        public List<double> Frequencies { get; private set; } = new List<double>();

        /// <summary>
        /// List of Responses for the corresponding frequencies for the test
        /// </summary>
        public List<double> Responses { get; private set; } = new List<double>();

        /// <summary>
        /// Name of the setup used for the test
        /// </summary>
        public string Setup { get; set; } = "";

        /// <summary>
        /// Caption used for the test
        /// </summary>
        public string Caption { get; set; } = "";

        /// <summary>
        /// Y-Axis label for the test
        /// </summary>
        public string Y_Axis { get; set; } = "";

        /// <summary>
        /// X-Axis label for the test
        /// </summary>
        public string X_Axis { get; set; } = "";

        /// <summary>
        /// Y-Offset used for the test
        /// </summary>
        public double Y_Offset { get; set; } = 0;

        /// <summary>
        /// Minimum frequency in Hz used for the test
        /// </summary>
        public double MinimumFrequency { get; set; } = 0;

        /// <summary>
        /// Maximum frequency in Hz used for the test
        /// </summary>
        public double MaximumFrequency { get; set; } = 0;

        /// <summary>
        /// The size of the step in Hz between minimum and maximum frequency
        /// </summary>
        public double StepSize { get; set; } = 0;

        /// <summary>
        /// The size of the output amplitude in Volts Pk-Pk
        /// </summary>
        public double OutputAmplitudeVPP { get; set; } = 0;

        /// <summary>
        /// The size of the output amplitude in RMS
        /// </summary>
        public double OutputAmplitudeRMS { get; set; } = 0;

        /// <summary>
        /// The channel that the data was acquired on
        /// </summary>
        public int AcquisitionChannel { get; set; } = 0;

        /// <summary>
        /// The date that the test was performed
        /// </summary>
        public DateTime Date { get; set; } = DateTime.Now;

        /// <summary>
        /// A unique id used to identify this object
        /// </summary>
        private int id { get; set; } = ++DatasetCount;

        /// <summary>
        /// A static count that indicates how many dataset objects are in existence,
        /// is used to set the id of each dataset object.
        /// </summary>
        public static int DatasetCount { get; private set; } = 0;

        #endregion

        #region Constructors

        /// <summary>
        /// Creates a dataset initialized with the given frequencies and responses
        /// </summary>
        /// <param name="setup">setup used</param>
        /// <param name="caption">setup caption used</param>
        /// <param name="y_axis">y-axis label used</param>
        /// <param name="x_axis">x-axis label used</param>
        /// <param name="y_offset">y-offset used</param>
        /// <param name="minFreq">minimum frequency (Hz)</param>
        /// <param name="maxFreq">maximum frequency (Hz)</param>
        /// <param name="stepSize">frequency step-size (Hz)</param>
        /// <param name="outputAmpVpp">output amplitude Vpp</param>
        /// <param name="outputAmpRMS">output amplitude RMS</param>
        /// <param name="channel">channel used</param>
        /// <param name="date">date of the test</param>
        /// <param name="frequencies">list of frequencies used</param>
        /// <param name="responses">list of responses used</param>
        public Dataset(string setup = "", string caption = "",
            string y_axis = "", string x_axis = "",
            double y_offset = 0, double minFreq = 0, double maxFreq = 0, double stepSize = 0,
            double outputAmpVpp = 0, double outputAmpRMS = 0,
            int channel = 0, DateTime date = new DateTime(),
            List<double> frequencies = null, List<double> responses = null)
        {
            Setup = setup;
            Caption = caption;
            Y_Axis = y_axis;
            X_Axis = x_axis;
            Y_Offset = y_offset;
            MinimumFrequency = minFreq;
            MaximumFrequency = maxFreq;
            OutputAmplitudeVPP = outputAmpVpp;
            OutputAmplitudeRMS = outputAmpRMS;
            AcquisitionChannel = channel;
            Date = date;
            Frequencies = frequencies ?? new List<double>();
            Responses = responses ?? new List<double>();
        }

        /// <summary>
        /// Creates a dataset initialized with the given frequencies and responses
        /// </summary>
        /// <param name="name">The unique name identifier given to this dataset, used to compare this dataset to others</param>
        /// <param name="setup">setup used</param>
        /// <param name="caption">setup caption used</param>
        /// <param name="y_axis">y-axis label used</param>
        /// <param name="x_axis">x-axis label used</param>
        /// <param name="y_offset">y-offset used</param>
        /// <param name="minFreq">minimum frequency (Hz)</param>
        /// <param name="maxFreq">maximum frequency (Hz)</param>
        /// <param name="stepSize">frequency step-size (Hz)</param>
        /// <param name="outputAmpVpp">output amplitude Vpp</param>
        /// <param name="outputAmpRMS">output amplitude RMS</param>
        /// <param name="channel">channel used</param>
        /// <param name="date">date of the test</param>
        /// <param name="frequencies">list of frequencies used</param>
        /// <param name="responses">list of responses used</param>
        public Dataset(string name, string setup = "", string caption = "",
            string y_axis = "", string x_axis = "",
            double y_offset = 0, double minFreq = 0, double maxFreq = 0, double stepSize = 0,
            double outputAmpVpp = 0, double outputAmpRMS = 0,
            int channel = 0, DateTime date = new DateTime(),
            List<double> frequencies = null, List<double> responses = null)
        {
            Name = name;
            Setup = setup;
            Caption = caption;
            Y_Axis = y_axis;
            X_Axis = x_axis;
            Y_Offset = y_offset;
            MinimumFrequency = minFreq;
            MaximumFrequency = maxFreq;
            OutputAmplitudeVPP = outputAmpVpp;
            OutputAmplitudeRMS = outputAmpRMS;
            AcquisitionChannel = channel;
            Date = date;
            Frequencies = frequencies ?? new List<double>();
            Responses = responses ?? new List<double>();
        }

        #endregion

        #region Methods

        /// <summary>
        /// Returns true if the given object is equal to the other, or if the given object is of type Dataset
        /// Returns true if the Names are the same.
        /// </summary>
        /// <param name="obj">Object to compare to</param>
        /// <returns>Returns true if the objects are equal, returns false otherwise</returns>
        public override bool Equals(object obj)
        {
            return base.Equals(obj) || 
                (obj.GetType() == GetType() && Name.Equals(((Dataset)obj).Name));
        }

        #region File IO

        /// <summary>
        /// Reads in a datatable from a file
        /// </summary>
        /// <param name="_name"></param>
        /// <param name="fileLocation"></param>
        /// <returns></returns>
        public static Dataset CreateDataTableFromFile(string _name, string fileLocation)
        {
            Dataset temp = new Dataset(name:fileLocation);

            StreamReader sr = new StreamReader(fileLocation);
            bool hitHeader = false;
            while (!sr.EndOfStream)
            {
                string input = sr.ReadLine();

                if (input == string.Empty)
                    continue;
                else
                {

                    string[] s = input.Split(new char[] { '\t' });

                    if (s.Length < 2)
                        throw new InvalidDataException("The data file must be formatted with tab-delimitted data containing frequency data and response data!");

                    // Reads the header info
                    if (s[0].StartsWith("Setup:"))
                        temp.Setup = s[1];
                    else if (s[0].StartsWith("Caption:"))
                        temp.Caption = s[1];
                    else if (s[0].StartsWith("Y-Axis:"))
                        temp.Y_Axis = s[1];
                    else if (s[0].StartsWith("X-Axis:"))
                        temp.X_Axis = s[1];
                    else if (s[0].StartsWith("Y-Offset"))
                        temp.Y_Offset = double.Parse(s[1]);
                    else if (s[0].StartsWith("Minimum Frequency"))
                        temp.MinimumFrequency = double.Parse(s[1]);
                    else if (s[0].StartsWith("Maximum Frequency"))
                        temp.MaximumFrequency = double.Parse(s[1]);
                    else if (s[0].StartsWith("Step Size"))
                        temp.StepSize = double.Parse(s[1]);
                    else if (s[0].StartsWith("Output Amplitude [Vpp]"))
                        temp.OutputAmplitudeVPP = double.Parse(s[1]);
                    else if (s[0].StartsWith("Output Amplitude [Vrms]"))
                        temp.OutputAmplitudeRMS = double.Parse(s[1]);
                    else if (s[0].StartsWith("Acquisition Channel"))
                        temp.AcquisitionChannel = int.Parse(s[1]);
                    else if (s[0].StartsWith("Date"))
                        temp.Date = DateTime.ParseExact(s[1], "dd.MM.yyyy", CultureInfo.InvariantCulture);
                    else if (s[0].StartsWith("Frequency"))
                        hitHeader = true;   // Determines when the parser has moved into the section of the file containing the frequency data
                    else if (hitHeader)
                    {
                        temp.Frequencies.Add(double.Parse(s[0]));
                        temp.Responses.Add(double.Parse(s[1]));
                    }
                    else
                        continue;
                }
            }
            sr.Close();

            return temp;
        }

        #endregion

        public virtual Excel.Range CreateData(Excel._Workbook workbook, int column_offset, int row_offset)
        {
            Excel._Worksheet data = workbook.Worksheets["Data"];

            // Gets the range from the current worksheet
            string columnLetter = ExcelPort.ColumnNumToColumnString(column_offset);
            string nextColumnLetter = ExcelPort.ColumnNumToColumnString(column_offset + 1);
            Excel.Range range = data.Range[columnLetter + row_offset, nextColumnLetter + (row_offset + 14 + Frequencies.Count)];

            
            // Writes the data to excel
            range.Value = GetStringData();
            return range;
        }

        #region CreateChart

        /// <summary>
        /// Creates an excel chart from the given data in the given workbook, at the given location.
        /// </summary>
        /// <param name="workbook">Workbook to create the chart in.</param>
        /// <param name="rData">The excel range containing the data for the chart</param>
        /// <param name="column_offset">Column to place the chart at in column number</param>
        /// <param name="row_offset">Row to place the chart at as a row number</param>
        /// <param name="height">[optional] Specifies the height in pixels of the chart (300)</param>
        /// <param name="width">[optional] Specifies the width in pixels of the chart (300)</param>
        /// <param name="graphTitle">[optional] Specifies the string title of the chart ("")</param>
        /// <param name="xAxis">[optional] Specifies the string x Axis title of the chart ("")</param>
        /// <param name="yAxis">[optional] Specifies the string y Axis title of the chart ("")</param>
        public virtual void CreateChart(Excel._Workbook workbook, Excel.Range rData, int column_offset, int row_offset,
            double height = 300, double width = 300, 
            string graphTitle = "", string xAxis = "", string yAxis = "")
        {
            Excel._Worksheet data = workbook.Worksheets["Data"];

            // Gets the range from the current worksheet
            string columnLetter = ExcelPort.ColumnNumToColumnString(column_offset);

            // Add chart.
            Excel.ChartObjects charts = data.ChartObjects();
            Excel.Range origin = data.Range[(columnLetter + row_offset)];
            Excel.ChartObject chartObject = charts.Add((double)origin.Top, (double)origin.Left, 300, 300);
            Excel.Chart chart = chartObject.Chart;

            // Sets the chart range
            Excel.Range range = CreateData(workbook, column_offset, row_offset);
            chart.SetSourceData(rData);

            // Set chart properties.
            chart.ChartType = Excel.XlChartType.xlLine;
            chart.ChartWizard(Source: rData,
                Title: graphTitle,
                CategoryTitle: xAxis,
                ValueTitle: yAxis);
        }

        /// <summary>
        /// Creates an excel chart from the given data in the given workbook, at the given location.
        /// </summary>
        /// <param name="workbook">Workbook to create the chart in.</param>
        /// <param name="data_column_offset">The column offset for the data for the chart</param>
        /// <param name="data_row_offset">The row offset for the data for the chart</param>
        /// <param name="column_offset">Column to place the chart at in column number</param>
        /// <param name="row_offset">Row to place the chart at as a row number</param>
        /// <param name="height">[optional] Specifies the height in pixels of the chart (300)</param>
        /// <param name="width">[optional] Specifies the width in pixels of the chart (300)</param>
        /// <param name="graphTitle">[optional] Specifies the string title of the chart ("")</param>
        /// <param name="xAxis">[optional] Specifies the string x Axis title of the chart ("")</param>
        /// <param name="yAxis">[optional] Specifies the string y Axis title of the chart ("")</param>
        public virtual void CreateChart(Excel._Workbook workbook, int data_column_offset, int data_row_offset, int column_offset, int row_offset,
            double height = 300, double width = 300,
            string graphTitle = "", string xAxis = "", string yAxis = "")
        {
            // Gets the range from the current worksheet
            string columnLetter = ExcelPort.ColumnNumToColumnString(data_column_offset);
            Excel.Range range = CreateData(workbook, data_column_offset, data_row_offset);

            // Calls the overloaded function
            CreateChart(workbook, range, column_offset, row_offset, height, width, graphTitle, xAxis, yAxis);
        }

        #endregion

        public virtual string[,] GetStringData()
        {
            string[,] dt = new string[14 + Frequencies.Count, 2];

            // ALL the header info.
            dt[0, 0] = "Frequency" + id;
            dt[0, 1] = "Response" + id;
            dt[1, 0] = Setup;
            dt[2, 0] = "Caption:";
            dt[2, 1] = Caption;
            dt[3, 0] = "Y-Axis";
            dt[3, 1] = Y_Axis;
            dt[4, 0] = "X-Axis";
            dt[4, 1] = X_Axis;
            dt[5, 0] = "Y-Offset";
            dt[5, 1] = Y_Offset.ToString();
            dt[6, 0] = "Minimum Frequency";
            dt[6, 1] = MinimumFrequency.ToString();
            dt[7, 0] = "Maximum Frequency";
            dt[7, 1] = MaximumFrequency.ToString();
            dt[8, 0] = "Step Size";
            dt[8, 1] = StepSize.ToString();
            dt[9, 0] = "Output Amplitude [Vpp]";
            dt[9, 1] = OutputAmplitudeVPP.ToString();
            dt[10, 0] = "Output Amplitude [RMS]";
            dt[10, 1] = OutputAmplitudeRMS.ToString();
            dt[11, 0] = "Acquisition Channel";
            dt[11, 1] = AcquisitionChannel.ToString();
            dt[12, 0] = "Date:";
            dt[12, 1] = Date.ToString();
            dt[13, 0] = "Frequency [Hz]";
            dt[13, 1] = "RMS [dB]";

            // Writes the frequency and response data down
            for (int i = 0; i < Frequencies.Count; i++)
            {
                dt[14 + i, 0] = Frequencies[i].ToString();
                dt[14 + i, 1] = Responses[i].ToString();
            }

            return dt;
        }

        public override string ToString()
        {
            return Name + Date.ToShortDateString();
        }

        /// <summary>
        /// Calculates the average of the response data
        /// </summary>
        /// <returns></returns>
        public double Average()
        {
            return Responses.Average();
        }

        /// <summary>
        /// Calculates the standard deviation of the response data
        /// std = sqrt(avg((x[i] - avg(x))^2))
        /// </summary>
        /// <returns></returns>
        public double StdDev()
        {
            double result = 0;
            double average = Average();
            foreach (double d in Responses)
                result += Math.Pow(d - average, 2);
            result /= Responses.Count;
            return Math.Sqrt(result);
        }

        public ICollection<Tuple<double, double>> getData()
        {
            List<Tuple<double, double>> result = new List<Tuple<double, double>>();
            for (int i = 0; i < Frequencies.Count; i++)
                result.Add(new Tuple<double, double>(Frequencies[i], Responses[i]));
            return result;
        }

        #endregion
    }
}
