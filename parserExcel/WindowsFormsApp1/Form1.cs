using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Cyriller;
using Cyriller.Model;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.Entity;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        Parser parser = new Parser();


        public Form1()
        {
            InitializeComponent();

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
                        parser.ProductItemProcessor(filename);
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
                    parser.ProductItemProcessor(dlg.FileName);
                });
            }

        }



        private void buttExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
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
                    var data = (from d in db.ProductItems
                                where d.Price > PriceMore &&
                                d.Price < PriceSmaller
                                select d);
                    dataGridView1.DataSource = data.ToList();
                }
                if (FilterName.Text.IsNullOrEmpty() && FilterCustomer.Text.IsNullOrEmpty() && FilterInvoice.Text.IsNullOrEmpty()
                    && monthCalendar1.SelectionStart != monthCalendar1.TodayDate)
                {
                    var data = (from d in db.ProductItems
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
                    var data = (from d in db.ProductItems
                                where d.Price > PriceMore &&
                                d.Price < PriceSmaller &&
                                d.PartNumber.Contains(PartNumber)
                                select d);
                    dataGridView1.DataSource = data.ToList();
                }
                else if (FilterName.Text.IsNotNullOrEmpty() && FilterCustomer.Text.IsNullOrEmpty() && FilterInvoice.Text.IsNullOrEmpty()
                    && monthCalendar1.SelectionStart != monthCalendar1.TodayDate)
                {
                    var data = (from d in db.ProductItems
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

                    var data = (from d in db.ProductItems
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

                    var data = (from d in db.ProductItems
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

                    var data = (from d in db.ProductItems
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
                    var data = (from d in db.ProductItems
                                where d.Price > PriceMore &&
                                d.Price < PriceSmaller &&

                                d.Customer.Contains(Customer)

                                select d);
                    dataGridView1.DataSource = data.ToList();
                }
                else if (FilterName.Text.IsNullOrEmpty() && FilterCustomer.Text.IsNotNullOrEmpty() && FilterInvoice.Text.IsNullOrEmpty()
                    && monthCalendar1.SelectionStart != monthCalendar1.TodayDate)
                {
                    var data = (from d in db.ProductItems
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

                    var data = (from d in db.ProductItems
                                where d.Price > PriceMore &&
                                d.Price < PriceSmaller &&

                                d.Customer.Contains(Customer) &&
                                d.Acct == Invoice

                                select d);
                    dataGridView1.DataSource = data.ToList();
                }
                else if (FilterName.Text.IsNotNullOrEmpty() && FilterCustomer.Text.IsNullOrEmpty() && FilterInvoice.Text.IsNotNullOrEmpty())
                {

                    var data = (from d in db.ProductItems
                                where d.Price > PriceMore &&
                                d.Price < PriceSmaller &&

                                d.PartNumber.Contains(PartNumber) &&
                                d.Acct == Invoice

                                select d);
                    dataGridView1.DataSource = data.ToList();

                }
                else if (FilterName.Text.IsNullOrEmpty() && FilterCustomer.Text.IsNullOrEmpty() && FilterInvoice.Text.IsNotNullOrEmpty())
                {

                    var data = (from d in db.ProductItems
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

            using (ProductItemContext db = new ProductItemContext())
            {
                string type = FilterType.Text;
                string Customer = FilterCustomerInvoice.Text;
                string Invoice = FilterInvoceInvoice.Text;

                if (FilterType.Text.IsNullOrEmpty() && FilterCustomerInvoice.Text.IsNullOrEmpty() && FilterInvoceInvoice.Text.IsNullOrEmpty()
                    && monthCalendar2.SelectionStart == monthCalendar2.TodayDate)
                {
                    var data = (from d in db.Invoices.Include(a => a.Customer)
                                where d.Sum > SumMore &&
                                d.Sum < SumSmaller
                                select d);
                    dataGridView2.DataSource = data.ToList().Select(a => new InvoicesAndCustomer
                    {
                        Acct = a.Acct,
                        Date = a.Date,
                        Id = a.Id,
                        Sum = a.Sum,
                        Type = a.Type,
                        CompanyName = a.Customer.CompanyName,
                        Inn = a.Customer.Inn
                    }).ToList();


                }
                if (FilterType.Text.IsNullOrEmpty() && FilterCustomerInvoice.Text.IsNullOrEmpty() && FilterInvoceInvoice.Text.IsNullOrEmpty()
                    && monthCalendar2.SelectionStart != monthCalendar2.TodayDate)
                {
                    var data = (from d in db.Invoices.Include(a => a.Customer)
                                where d.Sum > SumMore &&
                                d.Sum < SumSmaller &&
                                monthCalendar2.SelectionStart <= d.Date &&
                                monthCalendar2.SelectionEnd >= d.Date
                                select d);
                    dataGridView2.DataSource = data.ToList().Select(a => new InvoicesAndCustomer
                    {
                        Acct = a.Acct,
                        Date = a.Date,
                        Id = a.Id,
                        Sum = a.Sum,
                        Type = a.Type,
                        CompanyName = a.Customer.CompanyName,
                        Inn = a.Customer.Inn
                    }).ToList();
                }
                else if (FilterType.Text.IsNotNullOrEmpty() && FilterCustomerInvoice.Text.IsNullOrEmpty() && FilterInvoceInvoice.Text.IsNullOrEmpty()
                    && monthCalendar2.SelectionStart == monthCalendar2.TodayDate)
                {
                    var data = (from d in db.Invoices.Include(a => a.Customer)
                                where d.Sum > SumMore &&
                                d.Sum < SumSmaller &&
                                d.Type.Contains(type)
                                select d);
                    dataGridView2.DataSource = data.ToList().Select(a => new InvoicesAndCustomer
                    {
                        Acct = a.Acct,
                        Date = a.Date,
                        Id = a.Id,
                        Sum = a.Sum,
                        Type = a.Type,
                        CompanyName = a.Customer.CompanyName,
                        Inn = a.Customer.Inn
                    }).ToList();
                }
                else if (FilterType.Text.IsNotNullOrEmpty() && FilterCustomerInvoice.Text.IsNullOrEmpty() && FilterInvoceInvoice.Text.IsNullOrEmpty()
                    && monthCalendar2.SelectionStart != monthCalendar2.TodayDate)
                {
                    var data = (from d in db.Invoices.Include(a => a.Customer)
                                where d.Sum > SumMore &&
                                d.Sum < SumSmaller &&
                                d.Type.Contains(type) &&
                                monthCalendar2.SelectionStart <= d.Date &&
                                monthCalendar2.SelectionEnd >= d.Date
                                select d);
                    dataGridView2.DataSource = data.ToList().Select(a => new InvoicesAndCustomer
                    {
                        Acct = a.Acct,
                        Date = a.Date,
                        Id = a.Id,
                        Sum = a.Sum,
                        Type = a.Type,
                        CompanyName = a.Customer.CompanyName,
                        Inn = a.Customer.Inn
                    }).ToList();
                }
                else if (FilterType.Text.IsNotNullOrEmpty() && FilterCustomerInvoice.Text.IsNotNullOrEmpty() && FilterInvoceInvoice.Text.IsNullOrEmpty()
                    && monthCalendar2.SelectionStart == monthCalendar2.TodayDate)
                {

                    var data = (from d in db.Invoices.Include(a => a.Customer)
                                where d.Sum > SumMore &&
                                d.Sum < SumSmaller &&
                                d.Type.Contains(type) &&
                                d.Customer.CompanyName.Contains(Customer)
                                select d);
                    dataGridView2.DataSource = data.ToList().Select(a => new InvoicesAndCustomer
                    {
                        Acct = a.Acct,
                        Date = a.Date,
                        Id = a.Id,
                        Sum = a.Sum,
                        Type = a.Type,
                        CompanyName = a.Customer.CompanyName,
                        Inn = a.Customer.Inn
                    }).ToList();
                }
                else if (FilterType.Text.IsNotNullOrEmpty() && FilterCustomerInvoice.Text.IsNotNullOrEmpty() && FilterInvoceInvoice.Text.IsNullOrEmpty()
                    && monthCalendar2.SelectionStart != monthCalendar2.TodayDate)
                {

                    var data = (from d in db.Invoices.Include(a => a.Customer)
                                where d.Sum > SumMore &&
                                d.Sum < SumSmaller &&
                                d.Type.Contains(type) &&
                                d.Customer.CompanyName.Contains(Customer) &&
                                monthCalendar2.SelectionStart <= d.Date &&
                                monthCalendar2.SelectionEnd >= d.Date
                                select d);
                    dataGridView2.DataSource = data.ToList().Select(a => new InvoicesAndCustomer
                    {
                        Acct = a.Acct,
                        Date = a.Date,
                        Id = a.Id,
                        Sum = a.Sum,
                        Type = a.Type,
                        CompanyName = a.Customer.CompanyName,
                        Inn = a.Customer.Inn
                    }).ToList();
                }
                else if (FilterType.Text.IsNotNullOrEmpty() && FilterCustomerInvoice.Text.IsNotNullOrEmpty() && FilterInvoceInvoice.Text.IsNotNullOrEmpty())
                {

                    var data = (from d in db.Invoices.Include(a => a.Customer)
                                where d.Sum > SumMore &&
                                d.Sum < SumSmaller &&

                                d.Type.Contains(type) &&
                                d.Customer.CompanyName.Contains(Customer) &&
                                d.Acct == Invoice

                                select d);
                    dataGridView2.DataSource = data.ToList().Select(a => new InvoicesAndCustomer
                    {
                        Acct = a.Acct,
                        Date = a.Date,
                        Id = a.Id,
                        Sum = a.Sum,
                        Type = a.Type,
                        CompanyName = a.Customer.CompanyName,
                        Inn = a.Customer.Inn
                    }).ToList();
                }
                else if (FilterType.Text.IsNullOrEmpty() && FilterCustomerInvoice.Text.IsNotNullOrEmpty() && FilterInvoceInvoice.Text.IsNullOrEmpty()
                    && monthCalendar2.SelectionStart == monthCalendar2.TodayDate)
                {
                    var data = (from d in db.Invoices.Include(a => a.Customer)
                                where d.Sum > SumMore &&
                                d.Sum < SumSmaller &&

                                d.Customer.CompanyName.Contains(Customer)

                                select d);
                    dataGridView2.DataSource = data.ToList().Select(a => new InvoicesAndCustomer
                    {
                        Acct = a.Acct,
                        Date = a.Date,
                        Id = a.Id,
                        Sum = a.Sum,
                        Type = a.Type,
                        CompanyName = a.Customer.CompanyName,
                        Inn = a.Customer.Inn
                    }).ToList();
                }
                else if (FilterType.Text.IsNullOrEmpty() && FilterCustomerInvoice.Text.IsNotNullOrEmpty() && FilterInvoceInvoice.Text.IsNullOrEmpty()
                    && monthCalendar2.SelectionStart != monthCalendar2.TodayDate)
                {
                    var data = (from d in db.Invoices.Include(a => a.Customer)
                                where d.Sum > SumMore &&
                                d.Sum < SumSmaller &&

                                d.Customer.CompanyName.Contains(Customer) &&
                                monthCalendar1.SelectionStart <= d.Date &&
                                monthCalendar1.SelectionEnd >= d.Date

                                select d);
                    dataGridView2.DataSource = data.ToList().Select(a => new InvoicesAndCustomer
                    {
                        Acct = a.Acct,
                        Date = a.Date,
                        Id = a.Id,
                        Sum = a.Sum,
                        Type = a.Type,
                        CompanyName = a.Customer.CompanyName,
                        Inn = a.Customer.Inn
                    }).ToList();
                }
                else if (FilterType.Text.IsNullOrEmpty() && FilterCustomerInvoice.Text.IsNotNullOrEmpty() && FilterInvoceInvoice.Text.IsNotNullOrEmpty())
                {

                    var data = (from d in db.Invoices.Include(a => a.Customer)
                                where d.Sum > SumMore &&
                                d.Sum < SumSmaller &&

                                d.Customer.CompanyName.Contains(Customer) &&
                                d.Acct == Invoice

                                select d);
                    dataGridView2.DataSource = data.ToList().Select(a => new InvoicesAndCustomer
                    {
                        Acct = a.Acct,
                        Date = a.Date,
                        Id = a.Id,
                        Sum = a.Sum,
                        Type = a.Type,
                        CompanyName = a.Customer.CompanyName,
                        Inn = a.Customer.Inn
                    }).ToList();
                }
                else if (FilterType.Text.IsNotNullOrEmpty() && FilterCustomerInvoice.Text.IsNullOrEmpty() && FilterInvoceInvoice.Text.IsNotNullOrEmpty())
                {

                    var data = (from d in db.Invoices.Include(a => a.Customer)
                                where d.Sum > SumMore &&
                                d.Sum < SumSmaller &&

                                d.Type.Contains(type) &&
                                d.Acct == Invoice

                                select d);
                    dataGridView2.DataSource = data.ToList().Select(a => new InvoicesAndCustomer
                    {
                        Acct = a.Acct,
                        Date = a.Date,
                        Id = a.Id,
                        Sum = a.Sum,
                        Type = a.Type,
                        CompanyName = a.Customer.CompanyName,
                        Inn = a.Customer.Inn
                    }).ToList();
                }
                else if (FilterType.Text.IsNullOrEmpty() && FilterCustomerInvoice.Text.IsNullOrEmpty() && FilterInvoceInvoice.Text.IsNotNullOrEmpty())
                {

                    var data = (from d in db.Invoices.Include(a => a.Customer)
                                where d.Sum > SumMore &&
                                d.Sum < SumSmaller &&

                                d.Acct == Invoice

                                select d);
                    dataGridView2.DataSource = data.ToList().Select(a => new InvoicesAndCustomer
                    {
                        Acct = a.Acct,
                        Date = a.Date,
                        Id = a.Id,
                        Sum = a.Sum,
                        Type = a.Type,
                        CompanyName = a.Customer.CompanyName,
                        Inn = a.Customer.Inn
                    }).ToList();
                }

            }

        }

        private void buttFilterSearchCustomer_Click(object sender, EventArgs e)
        {
            string customerName = FilterCustomerName.Text;
            string inn = FilterCustomerINN.Text;
            string adress = FilterCustomerAdress.Text;
            using (ProductItemContext db = new ProductItemContext())
            {

                if (FilterCustomerName.Text.IsNullOrEmpty() && FilterCustomerINN.Text.IsNotNullOrEmpty() && FilterCustomerAdress.Text.IsNullOrEmpty())
                {
                    var data = (from d in db.Customers
                                where d.Inn.Contains(inn)
                                select d);
                    dataGridView3.DataSource = data.ToList();
                }

                if (FilterCustomerName.Text.IsNotNullOrEmpty() && FilterCustomerINN.Text.IsNullOrEmpty() && FilterCustomerAdress.Text.IsNullOrEmpty())
                {
                    var data = (from d in db.Customers
                                where d.CompanyName.Contains(customerName)
                                select d);
                    dataGridView3.DataSource = data.ToList();
                }
                if (FilterCustomerName.Text.IsNullOrEmpty() && FilterCustomerINN.Text.IsNullOrEmpty() && FilterCustomerAdress.Text.IsNotNullOrEmpty())
                {
                    var data = (from d in db.Customers
                                where d.Adress.Contains(adress)
                                select d);
                    dataGridView3.DataSource = data.ToList();
                }
                if (FilterCustomerName.Text.IsNullOrEmpty() && FilterCustomerINN.Text.IsNullOrEmpty() && FilterCustomerAdress.Text.IsNullOrEmpty())
                {
                    var data = (from d in db.Customers select d);
                    dataGridView3.DataSource = data.ToList();
                }
                if (FilterCustomerName.Text.IsNotNullOrEmpty() && FilterCustomerINN.Text.IsNotNullOrEmpty() && FilterCustomerAdress.Text.IsNullOrEmpty())
                {
                    var data = (from d in db.Customers
                                where d.Inn.Contains(inn) &&
                                d.CompanyName.Contains(customerName)
                                select d);
                    dataGridView3.DataSource = data.ToList();
                }
                if (FilterCustomerName.Text.IsNotNullOrEmpty() && FilterCustomerINN.Text.IsNotNullOrEmpty() && FilterCustomerAdress.Text.IsNotNullOrEmpty())
                {
                    var data = (from d in db.Customers
                                where d.Inn.Contains(inn) &&
                                d.CompanyName.Contains(customerName) &&
                                d.Adress.Contains(adress)
                                select d);
                    dataGridView3.DataSource = data.ToList();
                }
                if (FilterCustomerName.Text.IsNullOrEmpty() && FilterCustomerINN.Text.IsNotNullOrEmpty() && FilterCustomerAdress.Text.IsNotNullOrEmpty())
                {
                    var data = (from d in db.Customers
                                where d.Inn.Contains(inn) &&
                                d.Adress.Contains(adress)
                                select d);
                    dataGridView3.DataSource = data.ToList();
                }
                if (FilterCustomerName.Text.IsNotNullOrEmpty() && FilterCustomerINN.Text.IsNullOrEmpty() && FilterCustomerAdress.Text.IsNotNullOrEmpty())
                {
                    var data = (from d in db.Customers
                                where d.CompanyName.Contains(customerName) &&
                                d.Adress.Contains(adress)
                                select d);
                    dataGridView3.DataSource = data.ToList();
                }
            }
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

        private void butExportInCustomer_Click(object sender, EventArgs e)
        {
            Task task = Task.Factory.StartNew(() =>
            {
                Excel.Application exApp = new Excel.Application();
                exApp.Workbooks.Add();
                Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;

                for (int i = 0; i <= dataGridView3.Rows.Count - 1; i++)
                {
                    for (int j = 0; j <= dataGridView3.Columns.Count - 2; j++)
                    {
                        if (dataGridView3[j, i].Value == null)
                        {
                            wsh.Cells[i + 1, j + 1] = "";
                        }
                        else
                            wsh.Cells[i + 1, j + 1] = dataGridView3[j, i].Value.ToString();
                    }
                }
                wsh.Cells[dataGridView3.Rows.Count + 1, 1] = $"Количество покупателей {dataGridView3.Rows.Count}";
                exApp.Visible = true;
            });
        }

        private void butContrRF60_Click(object sender, EventArgs e)
        {
            string fileInvoice = "";
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Файлы эксель|*.xlsm|Все файлы|*.*";
            dlg.Title = "Выберите счет для спецификации";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                fileInvoice = dlg.FileName;
            }
            var helper = new wordHelper("ДоговорРФ60.docx");
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

        private void buttSpecRF100_Click(object sender, EventArgs e)
        {
            string fileInvoice = "";
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Файлы эксель|*.xlsm|Все файлы|*.*";
            dlg.Title = "Выберите счет для спецификации";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                fileInvoice = dlg.FileName;
            }
            var helper = new wordHelper("СпецификацияРФ100.docx");
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
                string FioSokr = helper.FioSokr(FIOSpecRFtextBox.Text);
                var items = new Dictionary<string, string>
            {
                {"{number}", NumberContractTextBox.Text  },
                {"{org}", CompanyNameSpecRFtextBox.Text  },
                {"{INN}", InnSpecRFtextBox.Text  },
                {"{KPP}", KppSpecRFtextBox.Text  },
                {"{Adress}", AdressSpecRFtextBox.Text  },
                {"{Bank}", BankSpecRFtextBox.Text  },
                {"{Bik}", BikSpecRFtextBox.Text  },
                {"{DATE}", DateContractSpecRFtextBox.Text  },
                {"{DATESpec}", dateTimePicker1.Value.ToString("dd.MM.yyyy")  },
                {"{dolg-im}", DolgnostSpecRFTextBox.Text  },
                {"{fioSokr}", FioSokr  },
                {"{r/s}", RSspecTextBox.Text },
                {"{k/s}", KSspecTextBox.Text  },
                {"{sumProp}", sumProp  },
                {"{sum}", sum.ToString("F" + 2)  },
                {"{nds}", nds.ToString("F" + 2)  },
            };
                helper.ProcessSpec(items, productItems);
            });
        }

        private void buttSpecRF60_Click(object sender, EventArgs e)
        {
            string fileInvoice = "";
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Файлы эксель|*.xlsm|Все файлы|*.*";
            dlg.Title = "Выберите счет для спецификации";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                fileInvoice = dlg.FileName;
            }
            var helper = new wordHelper("СпецификацияРФ60.docx");
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
                string FioSokr = helper.FioSokr(FIOSpecRFtextBox.Text);
                var items = new Dictionary<string, string>
            {
                {"{number}", NumberContractTextBox.Text  },
                {"{org}", CompanyNameSpecRFtextBox.Text  },
                {"{INN}", InnSpecRFtextBox.Text  },
                {"{KPP}", KppSpecRFtextBox.Text  },
                {"{Adress}", AdressSpecRFtextBox.Text  },
                {"{Bank}", BankSpecRFtextBox.Text  },
                {"{Bik}", BikSpecRFtextBox.Text  },
                {"{DATE}", DateContractSpecRFtextBox.Text  },
                {"{DATESpec}", dateTimePicker1.Value.ToString("dd.MM.yyyy")  },
                {"{dolg-im}", DolgnostSpecRFTextBox.Text  },
                {"{fioSokr}", FioSokr  },
                {"{r/s}", RSspecTextBox.Text },
                {"{k/s}", KSspecTextBox.Text  },
                {"{sumProp}", sumProp  },
                {"{sum}", sum.ToString("F" + 2)  },
                {"{nds}", nds.ToString("F" + 2)  },
            };
                helper.ProcessSpec(items, productItems);
            });
        }

        private void buttClearFormSpecRF_Click(object sender, EventArgs e)
        {
            NumberContractTextBox.Text = "";
            buttClearFormSpecRF.Text = "";
            CompanyNameSpecRFtextBox.Text = "";
            InnSpecRFtextBox.Text = "";
            KppSpecRFtextBox.Text = "";
            AdressSpecRFtextBox.Text = "";
            BankSpecRFtextBox.Text = "";
            BikSpecRFtextBox.Text = "";
            DateContractSpecRFtextBox.Text = "";
            DolgnostSpecRFTextBox.Text = "";
            RSspecTextBox.Text = "";
            KSspecTextBox.Text = "";
        }

        private void buttContractKZ100_Click(object sender, EventArgs e)
        {
            string fileInvoice = "";
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Файлы эксель|*.xlsm|Все файлы|*.*";
            dlg.Title = "Выберите счет для спецификации";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                fileInvoice = dlg.FileName;
            }
            var helper = new wordHelper("ДоговорКЗ.docx");
            var parserInvoice = new ParserInvoice(fileInvoice);
            List<ProductItem> productItems = new List<ProductItem>();
            Task task = Task.Factory.StartNew(() =>
            {
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
                CyrResult resultDolg = cyrPhrase.Decline(DolgnostContractKZtextBox.Text, GetConditionsEnum.Similar);
                CyrResult resultName = cyrName.Decline(FIOContractKZtextBox.Text);
                string FioSokr = helper.FioSokr(FIOContractKZtextBox.Text);
                string NumberContract = dateTimePicker1.Value.ToString("yyMdhm");
                var items = new Dictionary<string, string>
            {
                {"{number}", NumberContract  },
                {"{org}", CompanyNameContractKZtextBox.Text  },
                {"{dolg-rod}", resultDolg.Родительный  },
                {"{fio-rod}", resultName.Родительный  },
                {"{na-osnovanii}", NaOsnovaniiContractKZtextBox.Text  },
                {"{Bin}", BINContractKZtextBox.Text  },
                {"{Adress}", AdressContractKZtextBox.Text  },
                {"{Bank}", BankContractKZtextBox.Text  },
                {"{Bik}", BikContractKZtextBox.Text  },
                {"{DATE}", dateTimePicker1.Value.ToString("dd.MM.yyyy")  },
                {"{dolg-im}", DolgnostContractKZtextBox.Text  },
                {"{fio-im}", FIOContractKZtextBox.Text  },
                {"{fioSokr}", FioSokr  },
                {"{sumProp}", sumProp  },
                {"{sum}", sum.ToString("F" + 2)  },
            };
                helper.Process(items, productItems);
            });
        }

        private void buttContractKZ60_Click(object sender, EventArgs e)
        {
            string fileInvoice = "";
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Файлы эксель|*.xlsm|Все файлы|*.*";
            dlg.Title = "Выберите счет для спецификации";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                fileInvoice = dlg.FileName;
            }
            var helper = new wordHelper("ДоговорКЗ60.docx");
            var parserInvoice = new ParserInvoice(fileInvoice);
            List<ProductItem> productItems = new List<ProductItem>();
            Task task = Task.Factory.StartNew(() =>
            {
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
                CyrResult resultDolg = cyrPhrase.Decline(DolgnostContractKZtextBox.Text, GetConditionsEnum.Similar);
                CyrResult resultName = cyrName.Decline(FIOContractKZtextBox.Text);
                string FioSokr = helper.FioSokr(FIOContractKZtextBox.Text);
                string NumberContract = dateTimePicker1.Value.ToString("yyMdhm");
                var items = new Dictionary<string, string>
            {
                {"{number}", NumberContract  },
                {"{org}", CompanyNameContractKZtextBox.Text  },
                {"{dolg-rod}", resultDolg.Родительный  },
                {"{fio-rod}", resultName.Родительный  },
                {"{na-osnovanii}", NaOsnovaniiContractKZtextBox.Text  },
                {"{Bin}", BINContractKZtextBox.Text  },
                {"{Adress}", AdressContractKZtextBox.Text  },
                {"{Bank}", BankContractKZtextBox.Text  },
                {"{Bik}", BikContractKZtextBox.Text  },
                {"{DATE}", dateTimePicker1.Value.ToString("dd.MM.yyyy")  },
                {"{dolg-im}", DolgnostContractKZtextBox.Text  },
                {"{fio-im}", FIOContractKZtextBox.Text  },
                {"{fioSokr}", FioSokr  },
                {"{sumProp}", sumProp  },
                {"{sum}", sum.ToString("F" + 2)  },
            };
                helper.Process(items, productItems);
            });
        }

        private void buttClearFormKZ_Click(object sender, EventArgs e)
        {
            CompanyNameContractKZtextBox.Text = "";
            DolgnostContractKZtextBox.Text = "";
            NaOsnovaniiContractKZtextBox.Text = "";
            FIOContractKZtextBox.Text = "";
            AdressContractKZtextBox.Text = "";
            BINContractKZtextBox.Text = "";
            BankContractKZtextBox.Text = "";
            BikContractKZtextBox.Text = "";
        }

        private void buttSpecKZ100_Click(object sender, EventArgs e)
        {
            string fileInvoice = "";
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Файлы эксель|*.xlsm|Все файлы|*.*";
            dlg.Title = "Выберите счет для спецификации";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                fileInvoice = dlg.FileName;
            }
            var helper = new wordHelper("СпецификацияКЗ.docx");
            var parserInvoice = new ParserInvoice(fileInvoice);
            List<ProductItem> productItems = new List<ProductItem>();
            Task task = Task.Factory.StartNew(() =>
            {
                parserInvoice.ProductItemProcessor(productItems);
                double sum = 0;
                foreach (var item in productItems)
                {
                    sum += item.Sum;
                }
                string sumProp = Сумма.Пропись(sum, Валюта.Рубли);
                string FioSokr = helper.FioSokr(FioSpecKZtextBox.Text);
                var items = new Dictionary<string, string>
            {
                {"{number}", ContactNumberSpecKZtextBox.Text  },
                {"{org}", CompanyNameSpecKZtextBox.Text  },
                {"{Bin}", BinSpecKZtextBox.Text  },
                {"{Adress}", AdressSpecKZtextBox.Text  },
                {"{Bank}", BankSpecKZtextBox.Text  },
                {"{Bik}", BikSpecKZtextBox.Text  },
                {"{DATE}", ContracDatetextBox.Text  },
                {"{DATESpec}", dateTimePicker1.Value.ToString("dd.MM.yyyy")  },
                {"{dolg-im}", DolgnostSpecKZtextBox.Text  },
                {"{fioSokr}", FioSokr  },
                {"{sumProp}", sumProp  },
                {"{sum}", sum.ToString("F" + 2)  },
            };
                helper.ProcessSpec(items, productItems);
            });
        }

        private void buttSpecKZ60_Click(object sender, EventArgs e)
        {
            string fileInvoice = "";
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Файлы эксель|*.xlsm|Все файлы|*.*";
            dlg.Title = "Выберите счет для спецификации";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                fileInvoice = dlg.FileName;
            }
            var helper = new wordHelper("СпецификацияКЗ60.docx");
            var parserInvoice = new ParserInvoice(fileInvoice);
            List<ProductItem> productItems = new List<ProductItem>();
            Task task = Task.Factory.StartNew(() =>
            {
                parserInvoice.ProductItemProcessor(productItems);
                double sum = 0;
                foreach (var item in productItems)
                {
                    sum += item.Sum;
                }
                string sumProp = Сумма.Пропись(sum, Валюта.Рубли);
                string FioSokr = helper.FioSokr(FioSpecKZtextBox.Text);
                var items = new Dictionary<string, string>
            {
                {"{number}", ContactNumberSpecKZtextBox.Text  },
                {"{org}", CompanyNameSpecKZtextBox.Text  },
                {"{Bin}", BinSpecKZtextBox.Text  },
                {"{Adress}", AdressSpecKZtextBox.Text  },
                {"{Bank}", BankSpecKZtextBox.Text  },
                {"{Bik}", BikSpecKZtextBox.Text  },
                {"{DATE}", ContracDatetextBox.Text  },
                {"{DATESpec}", dateTimePicker1.Value.ToString("dd.MM.yyyy")  },
                {"{dolg-im}", DolgnostSpecKZtextBox.Text  },
                {"{fioSokr}", FioSokr  },
                {"{sumProp}", sumProp  },
                {"{sum}", sum.ToString("F" + 2)  },
            };
                helper.ProcessSpec(items, productItems);
            });
        }

        private void buttClearFormSpecKZ_Click(object sender, EventArgs e)
        {
            ContactNumberSpecKZtextBox.Text = "";
            ContracDatetextBox.Text = "";
            CompanyNameSpecKZtextBox.Text = "";
            DolgnostSpecKZtextBox.Text = "";
            FioSpecKZtextBox.Text = "";
            AdressSpecKZtextBox.Text = "";
            BinSpecKZtextBox.Text = "";
            BankSpecKZtextBox.Text = "";
            BikSpecKZtextBox.Text = "";
        }
    }
}
