using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace WindowsFormsApp1
{
    class CustomerInfo
    {
        public int Id { get; set; }
        public string CompanyName { get; set; }
        [Required]
        [StringLength(30)]
        [Index("Ix_Inn", IsUnique =true)]
        public string Inn { get; set; }
        public string Adress { get; set; }
        public string Tel { get; set; }
        public string Customer { get; set; }
    }
}
