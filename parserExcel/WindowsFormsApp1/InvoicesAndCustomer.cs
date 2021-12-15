using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ParserAndForms
{
    class InvoicesAndCustomer
    {
        public int Id { get; set; }
        //public string Company { get; set; }
        public double Sum { get; set; }
        public string Acct { get; set; }
        public DateTime Date { get; set; }
        public string Type { get; set; }

        public string CompanyName { get; set; }
        public string Inn { get; set; }
    }
}
