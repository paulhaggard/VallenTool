using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToolkit
{
    public static class Shared
    {
        /// <summary>
        /// Adds an element to a collection if it does not already exist in the collection
        /// </summary>
        /// <typeparam name="T">Type used in the collection</typeparam>
        /// <param name="list">list or collection to add the item to</param>
        /// <param name="item">item to add to the collection</param>
        public static void AddIfDNE<T>(ICollection<T> list, T item)
        {
            if (!list.Contains(item))
                list.Add(item);
        }

        /// <summary>
        /// Creates a new histogram given the data in the arr, and the number of bins specified, or using the bins specified
        /// </summary>
        /// <param name="arr">Data to create the histogram for</param>
        /// <param name="numBins">The number of bins to use</param>
        /// <param name="bins">[optional] The starting number for each bin</param>
        /// <returns></returns>
        public static List<int> createHistogram(IEnumerable<double> arr, double numBins, List<double> bins = null)
        {
            // Returns a histogram with no bins
            if (numBins == 0 && (bins == null || bins.Count == 0))
                return new List<int>();

            // Figures out if a bin list needs to be created, or if one was already provided
            bins = bins ?? new List<double>();
            if(bins.Count == 0)
            {
                double max = arr.Max();
                double min = arr.Min();
                double binSize = (max - min) / numBins;

                for (int i = 0; i < numBins; i++)
                    bins.Add(i * binSize);
            }
            numBins = bins.Count;

            // Initializes the bins array
            List<int> result = new List<int>();
            for (int i = 0; i < numBins; i++)
                result.Add(0);

            // Creates the histogram
            for(int i = 0; i < arr.Count(); i++)
            {
                for (int bin = 0; bin < numBins; bin++)
                    if (bin == numBins - 1 && arr.ElementAt(i) >= bins[bin])
                        result[bin]++;
                    else if (arr.ElementAt(i) >= bins[bin] && arr.ElementAt(i) < bins[bin + 1])
                        result[bin]++;
            }

            return result;
        }
    }
}
