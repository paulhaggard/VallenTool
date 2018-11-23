using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToolkit.DataFormatting
{
    public interface IDataManData<T>
    {
        ICollection<T> getData();
    }
}
