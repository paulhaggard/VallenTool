using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToolkit.DataFormatting
{
    /// <summary>
    /// A way to represent a sequence of data used in the Mr Plotter plotting function
    /// </summary>
    /// <typeparam name="T">Type of data to be plotted</typeparam>
    public interface IDataManData<T>
    {
        /// <summary>
        /// Gets a list of coordinate pairs in the format of (X, Y) tuples
        /// </summary>
        /// <returns></returns>
        ICollection<Tuple<T, T>> getData();
    }
}
