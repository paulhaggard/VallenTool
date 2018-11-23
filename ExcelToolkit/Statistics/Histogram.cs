using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToolkit.Statistics
{
    /// <summary>
    /// Embodies a histogram, used for the histogram wizard
    /// </summary>
    public class Histogram
    {
        #region Properties

        private int binCount = 0;
        /// <summary>
        /// Number of bins in the histogram
        /// </summary>
        public int BinCount { get => binCount; set => setBinCount(value); }

        /// <summary>
        /// Results of the bin calculation
        /// </summary>
        public List<int> BinData { get; private set; } = new List<int>();

        public List<List<double>> BinDataArray { get; private set; } = new List<List<double>>();

        private Dataset data = new Dataset();
        /// <summary>
        /// The dataset that's encoded into the histogram
        /// </summary>
        public Dataset Data { get => data; set => setData(value); }

        private ICollection<double> bins = new List<double>();
        /// <summary>
        /// The bins that are used to calculate the histogram
        /// </summary>
        public ICollection<double> Bins { get => bins; set => setBins(value); }

        /// <summary>
        /// A boolean flag specifying if the data in BinData is ready to be read yet, or if it's still being calculated
        /// True if the data is ready
        /// false otherwise
        /// </summary>
        public bool isDataReady { get; private set; } = false;

        #endregion

        #region Constructors

        /// <summary>
        /// Creates a new histogram with the given data
        /// If binCount is 0, and bins is null, then an empty histogram is made
        /// </summary>
        /// <param name="data">Data to be encoded into the histogram</param>
        /// <param name="binCount">[optional] Number of bins to use</param>
        /// <param name="bins">[optional] Manual specification for the bins, takes precedence over binCount</param>
        public Histogram(Dataset data, int binCount = 0, ICollection<double> bins = null)
        {
            BinCount = binCount;
            Data = data;
            Bins = bins ?? new List<double>();

            if (binCount != 0 && bins == null)
                GenerateBins();

            if (Bins.Count > 0)
                Calculate();
        }

        #endregion

        #region Methods

        #region Setters

        /// <summary>
        /// Sets the data as a new value
        /// </summary>
        /// <param name="data">new data to be encoded</param>
        /// <param name="regenerateBins">Determines if the bins should be regenerated using the new max and min present in the data</param>
        public void setData(Dataset data, bool regenerateBins = true)
        {
            this.data = data;

            if (regenerateBins)
                GenerateBins();

            Calculate();
        }

        /// <summary>
        /// Sets the bins to be used in calculation to a new value
        /// </summary>
        /// <param name="bins">new bins to be used in calculation</param>
        public void setBins(ICollection<double> bins)
        {
            this.bins = bins;
            binCount = bins.Count();
            Calculate();
        }

        /// <summary>
        /// Sets the bincount to a new value
        /// </summary>
        /// <param name="binCount">sets the bincount used to calculate the bins automatically to a new value</param>
        public void setBinCount(int binCount)
        {
            this.binCount = binCount;
            GenerateBins();
            Calculate();
        }

        #endregion

        /// <summary>
        /// Generates the bins used in calculation based off of the binCount
        /// </summary>
        private void GenerateBins()
        {
            bins = new List<double>(binCount);

            double max = data.Responses.Max();
            double min = data.Responses.Min();
            double binSize = (max - min) / binCount;

            for (int i = 0; i < binCount; i++)
                bins.Add(i * binSize);
        }

        /// <summary>
        /// Calculates the bin counts for each bin
        /// </summary>
        private void Calculate()
        {
            isDataReady = false;

            // Initializes the bins Array
            BinData = new List<int>();
            BinDataArray = new List<List<double>>();
            for (int i = 0; i < binCount; i++)
            {
                BinData.Add(0);
                BinDataArray.Add(new List<double>());
            }

            // Creates the histogram
            for (int i = 0; i < Data.Responses.Count(); i++)
            {
                for (int bin = 0; bin < binCount; bin++)
                    if (bin == binCount - 1 && Data.Responses.ElementAt(i) >= bins.ElementAt(bin))
                    {
                        BinData[bin]++;
                        BinDataArray[bin].Add(Data.Responses.ElementAt(i));
                    }
                    else if (Data.Responses.ElementAt(i) >= bins.ElementAt(bin) && Data.Responses.ElementAt(i) < bins.ElementAt(bin + 1))
                    {
                        BinData[bin]++;
                        BinDataArray[bin].Add(Data.Responses.ElementAt(i));
                    }
            }

            isDataReady = true;
        }

        #endregion
    }
}
