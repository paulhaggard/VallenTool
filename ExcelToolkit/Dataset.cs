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
    public class Dataset : IExcelData
    {
        #region Properties

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

        #endregion

        #region Methods

        #region File IO

        /// <summary>
        /// Reads in a datatable from a file
        /// </summary>
        /// <param name="_name"></param>
        /// <param name="fileLocation"></param>
        /// <returns></returns>
        public static Dataset CreateDataTableFromFile(string _name, string fileLocation)
        {
            Dataset temp = new Dataset();

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

        public void CreateData(Excel._Workbook workbook, int column_offset, int row_offset)
        {
            Excel._Worksheet data = workbook.Worksheets["Data"];

            // ALL the header info.
            data.Cells[row_offset, column_offset].Value = "Frequency" + id;
            data.Cells[row_offset, column_offset + 1].Value = "Response" + id;
            data.Cells[row_offset + 1, column_offset].Value = Setup;
            data.Cells[row_offset + 2, column_offset].Value = "Caption:";
            data.Cells[row_offset + 2, column_offset + 1].Value = Caption;
            data.Cells[row_offset + 3, column_offset].Value = "Y-Axis";
            data.Cells[row_offset + 3, column_offset + 1].Value = Y_Axis;
            data.Cells[row_offset + 4, column_offset].Value = "X-Axis";
            data.Cells[row_offset + 4, column_offset + 1].Value = X_Axis;
            data.Cells[row_offset + 5, column_offset].Value = "Y-Offset";
            data.Cells[row_offset + 5, column_offset + 1].Value = Y_Offset;
            data.Cells[row_offset + 6, column_offset].Value = "Minimum Frequency";
            data.Cells[row_offset + 6, column_offset + 1].Value = MinimumFrequency;
            data.Cells[row_offset + 7, column_offset].Value = "Maximum Frequency";
            data.Cells[row_offset + 7, column_offset + 1].Value = MaximumFrequency;
            data.Cells[row_offset + 8, column_offset].Value = "Step Size";
            data.Cells[row_offset + 8, column_offset + 1].Value = StepSize;
            data.Cells[row_offset + 9, column_offset].Value = "Output Amplitude [Vpp]";
            data.Cells[row_offset + 9, column_offset + 1].Value = OutputAmplitudeVPP;
            data.Cells[row_offset + 10, column_offset].Value = "Output Amplitude [RMS]";
            data.Cells[row_offset + 10, column_offset + 1].Value = OutputAmplitudeRMS;
            data.Cells[row_offset + 11, column_offset].Value = "Acquisition Channel";
            data.Cells[row_offset + 11, column_offset + 1].Value = AcquisitionChannel;
            data.Cells[row_offset + 12, column_offset].Value = "Date:";
            data.Cells[row_offset + 12, column_offset + 1].Value = Date;
            data.Cells[row_offset + 13, column_offset].Value = "Frequency [Hz]";
            data.Cells[row_offset + 13, column_offset + 1].Value = "RMS [dB]";

            // Writes the frequency and response data down
            for(int i = 0; i < Frequencies.Count; i++)
            {
                data.Cells[row_offset + 14 + i, column_offset].Value = Frequencies[i];
                data.Cells[row_offset + 14 + i, column_offset + 1].Value = Responses[i];
            }
        }

        public override string ToString()
        {
            return Caption.ToString() + Date.ToShortDateString();
        }

        #endregion
    }
}
