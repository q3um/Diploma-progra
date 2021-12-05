using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Cyriller;
using Cyriller.Model;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        List<ProductItem> items = new List<ProductItem>();
        BindingList<GridRow> data = new BindingList<GridRow>();
        Parser parser = new Parser();
        List<CustomerInfo> customerInfos = new List<CustomerInfo>();


        public Form1()
        {
            InitializeComponent();
        }

        private void Filtres_Enter(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void buttExport_Click(object sender, EventArgs e)
        {
            Task task = Task.Factory.StartNew(() =>
            {
                Excel.Application exApp = new Excel.Application();
                exApp.Workbooks.Add();
                Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;
                double sum = 0;

                for (int i = 0; i <= dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 0; j <= dataGridView1.Columns.Count - 2; j++)
                    {
                        if (dataGridView1[j, i].Value == null)
                        {
                            wsh.Cells[i + 1, j + 1] = "";
                        }
                        else
                            wsh.Cells[i + 1, j + 1] = dataGridView1[j, i].Value.ToString();
                    }
                    sum += Convert.ToDouble(dataGridView1[5, i].Value);
                }
                wsh.Cells[dataGridView1.Rows.Count + 1, 1] = $"Количество товаров {dataGridView1.Rows.Count}, на сумму {sum}";

                exApp.Visible = true;
            });
        }

        private void buttClearGrid_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
        }

        private void buttAddFolder_Click(object sender, EventArgs e)
        {
            {
                FolderBrowserDialog FBD = new FolderBrowserDialog();

                if (FBD.ShowDialog() == DialogResult.OK)
                {
                    MessageBox.Show(FBD.SelectedPath, "Вы выбрали", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    string[] fileList = Directory.GetFiles(FBD.SelectedPath, "*.xlsm", SearchOption.AllDirectories);
                    Stopwatch stopWatch = new Stopwatch();
                    stopWatch.Start();
                    //var filename = from n in fileList.AsParallel()
                    //               select n;
                    //foreach (var n in filename)
                    //{
                    //    if (n.Contains("~$"))
                    //    {
                    //        continue;
                    //    }
                    //    parser.ProductItemProcessor(n, items, customerInfos);
                    //}
                    //foreach (string filename in fileList)
                    //{
                    //    if (filename.Contains("~$"))
                    //    {
                    //        continue;
                    //    }
                    //    parser.ProductItemProcessor(filename, items, customerInfos);
                    //}

                    //List<Task> tasks = new List<Task>();
                    //foreach (string filename in fileList)
                    //{
                    //    if (filename.Contains("~$"))
                    //    {
                    //        continue;
                    //    }
                    //    var task = Task.Factory.StartNew(() => parser.ProductItemProcessor(filename, items, customerInfos));
                    //    tasks.Add(task);
                    //}
                    //Task.WaitAll(tasks.ToArray());

                    Parallel.ForEach(fileList, filename =>
                    {
                        if (filename.Contains("~$"))
                        {
                            return;
                        }
                        parser.ProductItemProcessor(filename, items, customerInfos);
                    });

                    stopWatch.Stop();
                    TimeSpan ts = stopWatch.Elapsed;
                    string elapsedTime = String.Format("{0:00}:{1:00}", ts.Minutes, ts.Seconds);
                    MessageBox.Show(elapsedTime);
                }
            }
        }
        private void buttAddFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Файлы эксель|*.xlsm|Все файлы|*.*";
            dlg.Multiselect = true;

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                if (dlg.FileName.Contains("~$"))
                {
                    return;
                }
                Task task = Task.Factory.StartNew(() =>
                {
                    parser.ProductItemProcessor(dlg.FileName, items, customerInfos);
                });
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void productItemBindingSource_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void buttExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void label38_Click(object sender, EventArgs e)
        {

        }

        private void label39_Click(object sender, EventArgs e)
        {

        }

        private void label43_Click(object sender, EventArgs e)
        {

        }

        private void label45_Click(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void butContrRF100_Click(object sender, EventArgs e)
        {
            string fileInvoice = "";
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Файлы эксель|*.xlsm|Все файлы|*.*";
            dlg.Title = "Выберите счет для спецификации";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                fileInvoice = dlg.FileName;
            }
            var helper = new wordHelper("Договор.docx");
            var parserInvoice = new ParserInvoice(fileInvoice);
            List<ProductItem> productItems = new List<ProductItem>();
            Task task = Task.Factory.StartNew(() =>
            {
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
            });

        }

        private void buttInvoice_Click(object sender, EventArgs e)
        {
            Process.Start(@"C:\элком_бланки\Blanks.xltm");
        }

        private void buttFilterSrch_Click(object sender, EventArgs e)
        {

            double PriceSmaller = default;
            if (FilterPriceSmaller.Text.IsNullOrEmpty())
                PriceSmaller = double.MaxValue;
            else
                PriceSmaller = Convert.ToDouble(FilterPriceSmaller.Text);

            double PriceMore = default;
            if (FilterPriceMore.Text.IsNullOrEmpty())
                PriceMore = 0;
            else
                PriceMore = Convert.ToDouble(FilterPriceMore.Text);

            using (ProductItemContext db = new ProductItemContext())
            {
                string PartNumber = FilterName.Text;
                string Customer = FilterCustomer.Text;
                string Invoice = FilterInvoice.Text;
                if (FilterName.Text.IsNullOrEmpty() && FilterCustomer.Text.IsNullOrEmpty() && FilterInvoice.Text.IsNullOrEmpty()
                    && monthCalendar1.SelectionStart == monthCalendar1.TodayDate)
                {
                    var data = (from d in db.productItems
                                where d.Price > PriceMore &&
                                d.Price < PriceSmaller
                                select d);
                    dataGridView1.DataSource = data.ToList();
                }
                if (FilterName.Text.IsNullOrEmpty() && FilterCustomer.Text.IsNullOrEmpty() && FilterInvoice.Text.IsNullOrEmpty()
                    && monthCalendar1.SelectionStart != monthCalendar1.TodayDate)
                {
                    var data = (from d in db.productItems
                                where d.Price > PriceMore &&
                                d.Price < PriceSmaller &&
                                monthCalendar1.SelectionStart <= d.Date &&
                                monthCalendar1.SelectionEnd >= d.Date
                                select d);
                    dataGridView1.DataSource = data.ToList();
                }
                else if (FilterName.Text.IsNotNullOrEmpty() && FilterCustomer.Text.IsNullOrEmpty() && FilterInvoice.Text.IsNullOrEmpty()
                    && monthCalendar1.SelectionStart == monthCalendar1.TodayDate)
                {
                    var data = (from d in db.productItems
                                where d.Price > PriceMore &&
                                d.Price < PriceSmaller &&
                                d.PartNumber.Contains(PartNumber)
                                select d);
                    dataGridView1.DataSource = data.ToList();
                }
                else if (FilterName.Text.IsNotNullOrEmpty() && FilterCustomer.Text.IsNullOrEmpty() && FilterInvoice.Text.IsNullOrEmpty()
                    && monthCalendar1.SelectionStart != monthCalendar1.TodayDate)
                {
                    var data = (from d in db.productItems
                                where d.Price > PriceMore &&
                                d.Price < PriceSmaller &&
                                d.PartNumber.Contains(PartNumber) &&
                                monthCalendar1.SelectionStart <= d.Date &&
                                monthCalendar1.SelectionEnd >= d.Date
                                select d);
                    dataGridView1.DataSource = data.ToList();
                }
                else if (FilterName.Text.IsNotNullOrEmpty() && FilterCustomer.Text.IsNotNullOrEmpty() && FilterInvoice.Text.IsNullOrEmpty()
                    && monthCalendar1.SelectionStart == monthCalendar1.TodayDate)
                {

                    var data = (from d in db.productItems
                                where d.Price > PriceMore &&
                                d.Price < PriceSmaller &&
                                d.PartNumber.Contains(PartNumber) &&
                                d.Customer.Contains(Customer)
                                select d);
                    dataGridView1.DataSource = data.ToList();
                }
                else if (FilterName.Text.IsNotNullOrEmpty() && FilterCustomer.Text.IsNotNullOrEmpty() && FilterInvoice.Text.IsNullOrEmpty()
                    && monthCalendar1.SelectionStart != monthCalendar1.TodayDate)
                {

                    var data = (from d in db.productItems
                                where d.Price > PriceMore &&
                                d.Price < PriceSmaller &&
                                d.PartNumber.Contains(PartNumber) &&
                                d.Customer.Contains(Customer) &&
                                monthCalendar1.SelectionStart <= d.Date &&
                                monthCalendar1.SelectionEnd >= d.Date
                                select d);
                    dataGridView1.DataSource = data.ToList();
                }
                else if (FilterName.Text.IsNotNullOrEmpty() && FilterCustomer.Text.IsNotNullOrEmpty() && FilterInvoice.Text.IsNotNullOrEmpty())
                {

                    var data = (from d in db.productItems
                                where d.Price > PriceMore &&
                                d.Price < PriceSmaller &&

                                d.PartNumber.Contains(PartNumber) &&
                                d.Customer.Contains(Customer) &&
                                d.Acct == Invoice

                                select d);
                    dataGridView1.DataSource = data.ToList();
                }
                else if (FilterName.Text.IsNullOrEmpty() && FilterCustomer.Text.IsNotNullOrEmpty() && FilterInvoice.Text.IsNullOrEmpty()
                    && monthCalendar1.SelectionStart == monthCalendar1.TodayDate)
                {
                    var data = (from d in db.productItems
                                where d.Price > PriceMore &&
                                d.Price < PriceSmaller &&

                                d.Customer.Contains(Customer)

                                select d);
                    dataGridView1.DataSource = data.ToList();
                }
                else if (FilterName.Text.IsNullOrEmpty() && FilterCustomer.Text.IsNotNullOrEmpty() && FilterInvoice.Text.IsNullOrEmpty()
                    && monthCalendar1.SelectionStart != monthCalendar1.TodayDate)
                {
                    var data = (from d in db.productItems
                                where d.Price > PriceMore &&
                                d.Price < PriceSmaller &&

                                d.Customer.Contains(Customer) &&
                                monthCalendar1.SelectionStart <= d.Date &&
                                monthCalendar1.SelectionEnd >= d.Date

                                select d);
                    dataGridView1.DataSource = data.ToList();
                }
                else if (FilterName.Text.IsNullOrEmpty() && FilterCustomer.Text.IsNotNullOrEmpty() && FilterInvoice.Text.IsNotNullOrEmpty())
                {

                    var data = (from d in db.productItems
                                where d.Price > PriceMore &&
                                d.Price < PriceSmaller &&

                                d.Customer.Contains(Customer) &&
                                d.Acct == Invoice

                                select d);
                    dataGridView1.DataSource = data.ToList();
                }
                else if (FilterName.Text.IsNotNullOrEmpty() && FilterCustomer.Text.IsNullOrEmpty() && FilterInvoice.Text.IsNotNullOrEmpty())
                {

                    var data = (from d in db.productItems
                                where d.Price > PriceMore &&
                                d.Price < PriceSmaller &&

                                d.PartNumber.Contains(PartNumber) &&
                                d.Acct == Invoice

                                select d);
                    dataGridView1.DataSource = data.ToList();
                }
                else if (FilterName.Text.IsNullOrEmpty() && FilterCustomer.Text.IsNullOrEmpty() && FilterInvoice.Text.IsNotNullOrEmpty())
                {

                    var data = (from d in db.productItems
                                where d.Price > PriceMore &&
                                d.Price < PriceSmaller &&

                                d.Acct == Invoice

                                select d);
                    dataGridView1.DataSource = data.ToList();
                }

            }


        }

        private void buttClearForm_Click(object sender, EventArgs e)
        {
            Organization.Text = "";
            Dolgnost.Text = "";
            NaOsnovanii.Text = "";
            FIO.Text = "";
            Adress.Text = "";
            INN.Text = "";
            KPP.Text = "";
            Bank.Text = "";
            BIK.Text = "";
            RS.Text = "";
            KS.Text = "";
        }

        private void FilterType_TextChanged(object sender, EventArgs e)
        {

        }

        private void FilterSerchInvoice_Click(object sender, EventArgs e)
        {

            double SumSmaller = default;
            if (FilterPriceSmaller.Text.IsNullOrEmpty())
                SumSmaller = double.MaxValue;
            else
                SumSmaller = Convert.ToDouble(FilterPriceSmaller.Text);

            double SumMore = default;
            if (FilterPriceMore.Text.IsNullOrEmpty())
                SumMore = 0;
            else
                SumMore = Convert.ToDouble(FilterPriceMore.Text);

            using (InvoiceContext db2 = new InvoiceContext())
            {
                string type = FilterType.Text;
                string Customer = FilterCustomerInvoice.Text;
                string Invoice = FilterInvoceInvoice.Text;
                if (FilterType.Text.IsNullOrEmpty() && FilterCustomerInvoice.Text.IsNullOrEmpty() && FilterInvoceInvoice.Text.IsNullOrEmpty()
                    && monthCalendar2.SelectionStart == monthCalendar2.TodayDate)
                {
                    var data = (from d in db2.invoices
                                where d.Sum > SumMore &&
                                d.Sum < SumSmaller
                                select d);
                    dataGridView2.DataSource = data.ToList();
                }
                if (FilterType.Text.IsNullOrEmpty() && FilterCustomerInvoice.Text.IsNullOrEmpty() && FilterInvoceInvoice.Text.IsNullOrEmpty()
                    && monthCalendar2.SelectionStart != monthCalendar2.TodayDate)
                {
                    var data = (from d in db2.invoices
                                where d.Sum > SumMore &&
                                d.Sum < SumSmaller &&
                                monthCalendar2.SelectionStart <= d.Date &&
                                monthCalendar2.SelectionEnd >= d.Date
                                select d);
                    dataGridView2.DataSource = data.ToList();
                }
                else if (FilterType.Text.IsNotNullOrEmpty() && FilterCustomerInvoice.Text.IsNullOrEmpty() && FilterInvoceInvoice.Text.IsNullOrEmpty()
                    && monthCalendar2.SelectionStart == monthCalendar2.TodayDate)
                {
                    var data = (from d in db2.invoices
                                where d.Sum > SumMore &&
                                d.Sum < SumSmaller &&
                                d.Type.Contains(type)
                                select d);
                    dataGridView2.DataSource = data.ToList();
                }
                else if (FilterType.Text.IsNotNullOrEmpty() && FilterCustomerInvoice.Text.IsNullOrEmpty() && FilterInvoceInvoice.Text.IsNullOrEmpty()
                    && monthCalendar2.SelectionStart != monthCalendar2.TodayDate)
                {
                    var data = (from d in db2.invoices
                                where d.Sum > SumMore &&
                                d.Sum < SumSmaller &&
                                d.Type.Contains(type) &&
                                monthCalendar2.SelectionStart <= d.Date &&
                                monthCalendar2.SelectionEnd >= d.Date
                                select d);
                    dataGridView2.DataSource = data.ToList();
                }
                else if (FilterType.Text.IsNotNullOrEmpty() && FilterCustomerInvoice.Text.IsNotNullOrEmpty() && FilterInvoceInvoice.Text.IsNullOrEmpty()
                    && monthCalendar2.SelectionStart == monthCalendar2.TodayDate)
                {

                    var data = (from d in db2.invoices
                                where d.Sum > SumMore &&
                                d.Sum < SumSmaller &&
                                d.Type.Contains(type) &&
                                d.Customer.Contains(Customer)
                                select d);
                    dataGridView2.DataSource = data.ToList();
                }
                else if (FilterType.Text.IsNotNullOrEmpty() && FilterCustomerInvoice.Text.IsNotNullOrEmpty() && FilterInvoceInvoice.Text.IsNullOrEmpty()
                    && monthCalendar2.SelectionStart != monthCalendar2.TodayDate)
                {

                    var data = (from d in db2.invoices
                                where d.Sum > SumMore &&
                                d.Sum < SumSmaller &&
                                d.Type.Contains(type) &&
                                d.Customer.Contains(Customer) &&
                                monthCalendar2.SelectionStart <= d.Date &&
                                monthCalendar2.SelectionEnd >= d.Date
                                select d);
                    dataGridView2.DataSource = data.ToList();
                }
                else if (FilterType.Text.IsNotNullOrEmpty() && FilterCustomerInvoice.Text.IsNotNullOrEmpty() && FilterInvoceInvoice.Text.IsNotNullOrEmpty())
                {

                    var data = (from d in db2.invoices
                                where d.Sum > SumMore &&
                                d.Sum < SumSmaller &&

                                d.Type.Contains(type) &&
                                d.Customer.Contains(Customer) &&
                                d.Acct == Invoice

                                select d);
                    dataGridView2.DataSource = data.ToList();
                }
                else if (FilterType.Text.IsNullOrEmpty() && FilterCustomerInvoice.Text.IsNotNullOrEmpty() && FilterInvoceInvoice.Text.IsNullOrEmpty()
                    && monthCalendar2.SelectionStart == monthCalendar2.TodayDate)
                {
                    var data = (from d in db2.invoices
                                where d.Sum > SumMore &&
                                d.Sum < SumSmaller &&

                                d.Customer.Contains(Customer)

                                select d);
                    dataGridView2.DataSource = data.ToList();
                }
                else if (FilterType.Text.IsNullOrEmpty() && FilterCustomerInvoice.Text.IsNotNullOrEmpty() && FilterInvoceInvoice.Text.IsNullOrEmpty()
                    && monthCalendar2.SelectionStart != monthCalendar2.TodayDate)
                {
                    var data = (from d in db2.invoices
                                where d.Sum > SumMore &&
                                d.Sum < SumSmaller &&

                                d.Customer.Contains(Customer) &&
                                monthCalendar1.SelectionStart <= d.Date &&
                                monthCalendar1.SelectionEnd >= d.Date

                                select d);
                    dataGridView2.DataSource = data.ToList();
                }
                else if (FilterType.Text.IsNullOrEmpty() && FilterCustomerInvoice.Text.IsNotNullOrEmpty() && FilterInvoceInvoice.Text.IsNotNullOrEmpty())
                {

                    var data = (from d in db2.invoices
                                where d.Sum > SumMore &&
                                d.Sum < SumSmaller &&

                                d.Customer.Contains(Customer) &&
                                d.Acct == Invoice

                                select d);
                    dataGridView2.DataSource = data.ToList();
                }
                else if (FilterType.Text.IsNotNullOrEmpty() && FilterCustomerInvoice.Text.IsNullOrEmpty() && FilterInvoceInvoice.Text.IsNotNullOrEmpty())
                {

                    var data = (from d in db2.invoices
                                where d.Sum > SumMore &&
                                d.Sum < SumSmaller &&

                                d.Type.Contains(type) &&
                                d.Acct == Invoice

                                select d);
                    dataGridView2.DataSource = data.ToList();
                }
                else if (FilterType.Text.IsNullOrEmpty() && FilterCustomerInvoice.Text.IsNullOrEmpty() && FilterInvoceInvoice.Text.IsNotNullOrEmpty())
                {

                    var data = (from d in db2.invoices
                                where d.Sum > SumMore &&
                                d.Sum < SumSmaller &&

                                d.Acct == Invoice

                                select d);
                    dataGridView2.DataSource = data.ToList();
                }

            }

        }

        private void buttFilterSearchCustomer_Click(object sender, EventArgs e)
        {
            
                string customerName = FilterCustomerName.Text;
                string inn = FilterCustomerINN.Text;
                string adress = FilterCustomerAdress.Text;
                using (CustomerInfoContext db3 = new CustomerInfoContext())
                {

                    if (FilterCustomerName.Text.IsNullOrEmpty() && FilterCustomerINN.Text.IsNotNullOrEmpty() && FilterCustomerAdress.Text.IsNullOrEmpty())
                    {
                        var data = (from d in db3.customerInfos
                                    where d.Inn.Contains(inn)
                                    select d);
                        dataGridView3.DataSource = data.ToList();
                    }
                    if (FilterCustomerName.Text.IsNotNullOrEmpty() && FilterCustomerINN.Text.IsNullOrEmpty() && FilterCustomerAdress.Text.IsNullOrEmpty())
                    {
                        var data = (from d in db3.customerInfos
                                    where d.CompanyName.Contains(customerName)
                                    select d);
                        dataGridView3.DataSource = data.ToList();
                    }
                    if (FilterCustomerName.Text.IsNullOrEmpty() && FilterCustomerINN.Text.IsNullOrEmpty() && FilterCustomerAdress.Text.IsNotNullOrEmpty())
                    {
                        var data = (from d in db3.customerInfos
                                    where d.Adress.Contains(adress)
                                    select d);
                        dataGridView3.DataSource = data.ToList();
                    }
                    if (FilterCustomerName.Text.IsNullOrEmpty() && FilterCustomerINN.Text.IsNullOrEmpty() && FilterCustomerAdress.Text.IsNullOrEmpty())
                    {
                        var data = (from d in db3.customerInfos select d);
                        dataGridView3.DataSource = data.ToList();
                    }
                    if (FilterCustomerName.Text.IsNotNullOrEmpty() && FilterCustomerINN.Text.IsNotNullOrEmpty() && FilterCustomerAdress.Text.IsNullOrEmpty())
                    {
                        var data = (from d in db3.customerInfos
                                    where d.Inn.Contains(inn) &&
                                    d.CompanyName.Contains(customerName)
                                    select d);
                        dataGridView3.DataSource = data.ToList();
                    }
                    if (FilterCustomerName.Text.IsNotNullOrEmpty() && FilterCustomerINN.Text.IsNotNullOrEmpty() && FilterCustomerAdress.Text.IsNotNullOrEmpty())
                    {
                        var data = (from d in db3.customerInfos
                                    where d.Inn.Contains(inn) &&
                                    d.CompanyName.Contains(customerName) &&
                                    d.Adress.Contains(adress)
                                    select d);
                        dataGridView3.DataSource = data.ToList();
                    }
                    if (FilterCustomerName.Text.IsNullOrEmpty() && FilterCustomerINN.Text.IsNotNullOrEmpty() && FilterCustomerAdress.Text.IsNotNullOrEmpty())
                    {
                        var data = (from d in db3.customerInfos
                                    where d.Inn.Contains(inn) &&
                                    d.Adress.Contains(adress)
                                    select d);
                        dataGridView3.DataSource = data.ToList();
                    }
                    if (FilterCustomerName.Text.IsNotNullOrEmpty() && FilterCustomerINN.Text.IsNullOrEmpty() && FilterCustomerAdress.Text.IsNotNullOrEmpty())
                    {
                        var data = (from d in db3.customerInfos
                                    where d.CompanyName.Contains(customerName) &&
                                    d.Adress.Contains(adress)
                                    select d);
                        dataGridView3.DataSource = data.ToList();
                    }
                }
            
        }

        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {

        }

        private void buttExportInvoice_Click(object sender, EventArgs e)
        {
            Task task = Task.Factory.StartNew(() =>
            {
                Excel.Application exApp = new Excel.Application();
                exApp.Workbooks.Add();
                Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;
                double sum = 0;

                for (int i = 0; i <= dataGridView2.Rows.Count - 1; i++)
                {
                    for (int j = 0; j <= dataGridView2.Columns.Count - 2; j++)
                    {
                        if (dataGridView2[j, i].Value == null)
                        {
                            wsh.Cells[i + 1, j + 1] = "";
                        }
                        else
                            wsh.Cells[i + 1, j + 1] = dataGridView2[j, i].Value.ToString();
                    }
                    sum += Convert.ToDouble(dataGridView2[2, i].Value);
                }
                wsh.Cells[dataGridView2.Rows.Count + 1, 1] = $"Количество счетов {dataGridView2.Rows.Count}, на сумму {sum}";
                exApp.Visible = true;
            });
        }
    }
}
