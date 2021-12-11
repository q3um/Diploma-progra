using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    class ProductItemContext : DbContext
    {
        public ProductItemContext() : base(nameOrConnectionString: "DBConnection") { }

        public DbSet<ProductItem> ProductItems { get; set; }
        public DbSet<Invoice> Invoices { get; set; }
        public DbSet<Customer> Customers { get; set; }
        public DbSet<InvoicesAndCustomer> InvoicesAndCustomers { get; set; }
    }
}
