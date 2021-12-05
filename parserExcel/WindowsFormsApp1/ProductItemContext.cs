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

        public DbSet<ProductItem> productItems { get; set; }
        public DbSet<Invoice> invoices { get; set; }
        public DbSet<CustomerInfo> customerInfos { get; set; }
    }
}
