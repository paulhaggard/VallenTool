using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Emerson_Excel_Tool
{
    public class FileStats
    {

        public string FileName { get; set; }
        public string FileShortPath { get; set; }
        public string FileFullPath { get; set; }
        public string FileDate { get; set; }
        public string Flagged { get; set; }
        public bool IsFlagged { get; set; } = false;
        public bool IsFavourite { get; set; }
        public string[] Tags { get; set; }

    }
}
