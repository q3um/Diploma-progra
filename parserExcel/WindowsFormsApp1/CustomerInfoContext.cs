using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    class CustomerInfoContext : DbContext
    {
        public CustomerInfoContext() : base("DbConnection") { }
        public DbSet<CustomerInfo> customerInfos { get; set; }
    }
}
