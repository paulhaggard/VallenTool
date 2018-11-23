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
    }
}
