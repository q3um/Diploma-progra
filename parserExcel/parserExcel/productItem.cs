using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace parserExcel
{
    class ProductItem
    {
        public string Customer { get; set; }
        public string PartNumber { get; set; }
        public int Quanity { get; set; }
        public double Price { get; set; }
        public double Sum { get; set; }
        public long Acct { get; set; }
        public string Date { get; set; }
        public string Type { get; set; }
    }
}
