using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToolkit.DataFormatting
{
    public static class Manip
    {
        /// <summary>
        /// Creates a list of datasets that contains all of the data for a small range of frequencies
        /// </summary>
        /// <param name="data">The data to put into the bins</param>
        /// <param name="numBins">The number of bins to use</param>
        /// <returns></returns>
        public static List<Dataset> GenerateFrequencyBins(ICollection<Dataset> data, int numBins)
        {
            List<Dataset> result = new List<Dataset>(numBins);
            for (int i = 0; i < numBins; i++)
                result.Add(new Dataset());

            double min = 2147483647;
            double max = -2147483648;
            foreach(Dataset set in data)
            {
                if (set.Responses.Min() < min)
                    min = set.Frequencies.Min();
                if (set.Responses.Max() > max)
                    max = set.Frequencies.Max();
            }
            double interval = max - min;

            foreach(Dataset set in data)
            {
                for(int i = 0; i < set.Frequencies.Count; i++)
                {
                    for(int k = 0; k < numBins; k++)
                    {
                        if (set.Frequencies[i] >= (k * interval) + min && k == numBins - 1)
                            result[k].Frequencies.Add(set.Frequencies[i]);
                        else if (set.Frequencies[i] >= (k * interval) + min && set.Frequencies[i] < ((k + 1) * interval) + min)
                            result[k].Frequencies.Add(set.Frequencies[i]);
                    }
                }
            }

            foreach(Dataset set in data)
            {
                foreach(Tuple<double, double> coord in set.getData())
                {
                    foreach(Dataset bin in result)
                    {
                        if (bin.Frequencies.Contains(coord.Item1))
                            bin.Responses.Add(coord.Item2);
                    }
                }
            }

            return result;
        }
    }
}
