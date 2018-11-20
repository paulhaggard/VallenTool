using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToolkit
{
    public static class Shared
    {
        public static void AddIfDNE<T>(List<T> list, T item)
        {
            if (!list.Contains(item))
                list.Add(item);
        }
    }
}
