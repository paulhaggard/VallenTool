using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToolkit
{
    /// <summary>
    /// A class used to contain a data file of frequencies and response data.
    /// </summary>
    public class Dataset
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

        #endregion
    }
}
