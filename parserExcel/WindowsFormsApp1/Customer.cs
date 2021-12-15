using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace ParserAndForms
{
    class Customer
    {
        public int Id { get; set; }
        public string CompanyName { get; set; }
        public string Inn { get; set; }
        public string Adress { get; set; }
        public string Tel { get; set; }
        public string CustomerFull { get; set; }

        public virtual ICollection<Invoice> Invoices { get; set; }

        public override string ToString()
        {
            return CompanyName;
        }
    }
}
