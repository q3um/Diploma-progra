using Cyriller;
using Cyriller.Model;
using System;
using System.Collections.Generic;
using System.Windows.Forms;


namespace ContractFill
{
    public partial class ContractFill : Form
    {
        public ContractFill()
        {
            InitializeComponent();
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            var helper = new wordHelper("Договор.docx");
            var parserInvoice = new ParserInvoice("Счет.xlsm");
            List<ProductItem> productItems = new List<ProductItem>();
            parserInvoice.ProductItemProcessor(productItems);
            double sum = 0;
            double nds = 0;
            foreach (var item in productItems)
            {
                sum += item.Sum;
            }
            nds = sum * 20 / 120;
            string sumProp = Сумма.Пропись(sum, Валюта.Рубли);
            CyrNounCollection cyrNounCollection = new CyrNounCollection();
            CyrAdjectiveCollection cyrAdjectiveCollection = new CyrAdjectiveCollection();
            CyrPhrase cyrPhrase = new CyrPhrase(cyrNounCollection, cyrAdjectiveCollection);
            CyrName cyrName = new CyrName();
            CyrResult resultDolg = cyrPhrase.Decline(Dolgnost.Text, GetConditionsEnum.Similar);
            CyrResult resultName = cyrName.Decline(FIO.Text);
            string FioSokr = helper.FioSokr(FIO.Text);
            string NumberContract = dateTimePicker1.Value.ToString("yyMdhm");
            var items = new Dictionary<string, string>
            {
                {"{number}", NumberContract  },
                {"{org}", Organization.Text  },
                {"{dolg-rod}", resultDolg.Родительный  },
                {"{fio-rod}", resultName.Родительный  },
                {"{na-osnovanii}", NaOsnovanii.Text  },
                {"{INN}", INN.Text  },
                {"{KPP}", KPP.Text  },
                {"{Adress}", Adress.Text  },
                {"{Bank}", Bank.Text  },
                {"{Bik}", BIK.Text  },
                {"{DATE}", dateTimePicker1.Value.ToString("dd.MM.yyyy")  },
                {"{dolg-im}", Dolgnost.Text  },
                {"{fio-im}", FIO.Text  },
                {"{fioSokr}", FioSokr  },
                {"{r/s}", RS.Text },
                {"{k/s}", KS.Text  },
                {"{sumProp}", sumProp  },
                {"{sum}", sum.ToString("F" + 2)  },
                {"{nds}", nds.ToString("F" + 2)  },

            };

            helper.Process(items, productItems);
        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        //private void button3_Click(object sender, EventArgs e)
        //{
            
        //}

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            var helper = new wordHelper("ДоговорКЗ.docx");
            var parserInvoice = new ParserInvoice("Счет.xlsm");
            List<ProductItem> productItems = new List<ProductItem>();
            parserInvoice.ProductItemProcessor(productItems);
            double sum = 0;
            foreach (var item in productItems)
            {
                sum += item.Sum;
            }
            string sumProp = Сумма.Пропись(sum, Валюта.Рубли);
            CyrNounCollection cyrNounCollection = new CyrNounCollection();
            CyrAdjectiveCollection cyrAdjectiveCollection = new CyrAdjectiveCollection();
            CyrPhrase cyrPhrase = new CyrPhrase(cyrNounCollection, cyrAdjectiveCollection);
            CyrName cyrName = new CyrName();
            CyrResult resultDolg = cyrPhrase.Decline(Dolgnost.Text, GetConditionsEnum.Similar);
            CyrResult resultName = cyrName.Decline(FIO.Text);
            string FioSokr = helper.FioSokr(FIO.Text);
            string NumberContract = dateTimePicker1.Value.ToString("yyMdhm");
            var items = new Dictionary<string, string>
            {
                {"{number}", NumberContract  },
                {"{org}", Organization.Text  },
                {"{dolg-rod}", resultDolg.Родительный  },
                {"{fio-rod}", resultName.Родительный  },
                {"{na-osnovanii}", NaOsnovanii.Text  },
                {"{Adress}", Adress.Text  },
                {"{Bank}", Bank.Text  },
                {"{Bin}", BIN.Text  },
                {"{Bik}", BIK.Text  },
                {"{DATE}", dateTimePicker1.Value.ToString("dd.MM.yyyy")  },
                {"{dolg-im}", Dolgnost.Text  },
                {"{fio-im}", FIO.Text  },
                {"{fioSokr}", FioSokr  },
                {"{sumProp}", sumProp  },
                {"{sum}", sum.ToString("F" + 2)  },

            };

            helper.Process(items, productItems);
        }

        private void tableLayoutPanel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            var helper = new wordHelper("ДоговорРФ60.docx");
            var parserInvoice = new ParserInvoice("Счет.xlsm");
            List<ProductItem> productItems = new List<ProductItem>();
            parserInvoice.ProductItemProcessor(productItems);
            double sum = 0;
            double nds = 0;
            foreach (var item in productItems)
            {
                sum += item.Sum;
            }
            nds = sum * 20 / 120;
            string sumProp = Сумма.Пропись(sum, Валюта.Рубли);
            CyrNounCollection cyrNounCollection = new CyrNounCollection();
            CyrAdjectiveCollection cyrAdjectiveCollection = new CyrAdjectiveCollection();
            CyrPhrase cyrPhrase = new CyrPhrase(cyrNounCollection, cyrAdjectiveCollection);
            CyrName cyrName = new CyrName();
            CyrResult resultDolg = cyrPhrase.Decline(Dolgnost.Text, GetConditionsEnum.Similar);
            CyrResult resultName = cyrName.Decline(FIO.Text);
            string FioSokr = helper.FioSokr(FIO.Text);
            string NumberContract = dateTimePicker1.Value.ToString("yyMdhm");
            var items = new Dictionary<string, string>
            {
                {"{number}", NumberContract  },
                {"{org}", Organization.Text  },
                {"{dolg-rod}", resultDolg.Родительный  },
                {"{fio-rod}", resultName.Родительный  },
                {"{na-osnovanii}", NaOsnovanii.Text  },
                {"{INN}", INN.Text  },
                {"{KPP}", KPP.Text  },
                {"{Adress}", Adress.Text  },
                {"{Bank}", Bank.Text  },
                {"{Bik}", BIK.Text  },
                {"{DATE}", dateTimePicker1.Value.ToString("dd.MM.yyyy")  },
                {"{dolg-im}", Dolgnost.Text  },
                {"{fio-im}", FIO.Text  },
                {"{fioSokr}", FioSokr  },
                {"{r/s}", RS.Text },
                {"{k/s}", KS.Text  },
                {"{sumProp}", sumProp  },
                {"{sum}", sum.ToString("F" + 2)  },
                {"{nds}", nds.ToString("F" + 2)  },

            };

            helper.Process(items, productItems);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var helper = new wordHelper("ДоговорКЗ60.docx");
            var parserInvoice = new ParserInvoice("Счет.xlsm");
            List<ProductItem> productItems = new List<ProductItem>();
            parserInvoice.ProductItemProcessor(productItems);
            double sum = 0;
            foreach (var item in productItems)
            {
                sum += item.Sum;
            }
            string sumProp = Сумма.Пропись(sum, Валюта.Рубли);
            CyrNounCollection cyrNounCollection = new CyrNounCollection();
            CyrAdjectiveCollection cyrAdjectiveCollection = new CyrAdjectiveCollection();
            CyrPhrase cyrPhrase = new CyrPhrase(cyrNounCollection, cyrAdjectiveCollection);
            CyrName cyrName = new CyrName();
            CyrResult resultDolg = cyrPhrase.Decline(Dolgnost.Text, GetConditionsEnum.Similar);
            CyrResult resultName = cyrName.Decline(FIO.Text);
            string FioSokr = helper.FioSokr(FIO.Text);
            string NumberContract = dateTimePicker1.Value.ToString("yyMdhm");
            var items = new Dictionary<string, string>
            {
                {"{number}", NumberContract  },
                {"{org}", Organization.Text  },
                {"{dolg-rod}", resultDolg.Родительный  },
                {"{fio-rod}", resultName.Родительный  },
                {"{na-osnovanii}", NaOsnovanii.Text  },
                {"{Adress}", Adress.Text  },
                {"{Bank}", Bank.Text  },
                {"{Bin}", BIN.Text  },
                {"{Bik}", BIK.Text  },
                {"{DATE}", dateTimePicker1.Value.ToString("dd.MM.yyyy")  },
                {"{dolg-im}", Dolgnost.Text  },
                {"{fio-im}", FIO.Text  },
                {"{fioSokr}", FioSokr  },
                {"{sumProp}", sumProp  },
                {"{sum}", sum.ToString("F" + 2)  },

            };

            helper.Process(items, productItems);
        }
    }
}
