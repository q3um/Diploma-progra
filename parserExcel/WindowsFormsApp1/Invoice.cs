using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    class Invoice
    {
        public int Id { get; set; }
        public string Customer { get; set; }
        public double Sum { get; set; }
        public string Acct { get; set; }
        public DateTime Date { get; set; }
        public string Type { get; set; }
    }
}
