using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ParserAndForms
{
    class GridRow
    {
        public string customer { get; set; } //обязательно нужно использовать get конструкцию
        public string partNumber { get; set; }
        public int quanity { get; set; }
        public double price { get; set; }
        public double sum { get; set; }
        public string acct { get; set; }
        public DateTime date { get; set; }
        public string type { get; set; }

        public string Hidden = ""; //Данное свойство не будет отображаться как колонка

        public GridRow(string customer, string partNumber, int quanity, double price, double sum, string acct, DateTime date, string type)
        {
            this.customer = customer;
            this.partNumber = partNumber;
            this.quanity = quanity;
            this.price = price;
            this.sum = sum;
            this.acct = acct;
            this.date = date;
            this.type = type;
        }
    }
}
