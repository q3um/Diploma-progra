using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    class InvoiceContext : DbContext
    {
            public InvoiceContext() : base("DbConnection") { }

            public DbSet<Invoice> invoices { get; set; }
    }
}
