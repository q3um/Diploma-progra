using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.Entity;
using System.Windows.Forms;

namespace ParserAndForms
{
    class Parser
    {
        public void SaveData(ProductItem productItem, ref double sum, ref string customer, ref string acct, ref DateTime date, ref string type)
        {
            productItem.Sum = productItem.Price * productItem.Quanity;
            sum += productItem.Sum;
            customer = productItem.Customer;
            acct = productItem.Acct;
            date = productItem.Date;
            type = productItem.Type;
        }
        public void ProductItemProcessor(string fileName)
        {
            List<ProductItem> items = new List<ProductItem>();

            using (ProductItemContext db = new ProductItemContext())
            {
                    Customer customerInfo = new Customer();
                    double sum = 0;
                    string acct = string.Empty;
                    string customer = string.Empty;
                    DateTime date = default;
                    string type = string.Empty;
                    Excel.Application excelApp = new Excel.Application();
                    Excel.Workbook excelWorkBook;
                    Excel.Worksheet excelSheet;
                    excelWorkBook = excelApp.Workbooks.Open(fileName);
                try
                {
                    int countNameInKP = 15; // строка с которой начинаются товары в кп
                    string nameList = (excelApp.Sheets[1] as Excel.Worksheet).Name;
                    excelSheet = excelWorkBook.Worksheets[nameList];
                    Regex regexAcctInvoice = new Regex(RegularFormular.AcctInInvoice);
                    Regex regexDate = new Regex(RegularFormular.DateInKPSheet);
                    Regex regexAcct = new Regex(RegularFormular.AcctInKpInvoice);
                    if (nameList == "НДС внутри" || nameList == "НДС сверху" || nameList == "НДС 0%")
                    {
                        //if (excelSheet.Range["B16"].Value == null & excelSheet.Range["B5"].Value == null)
                        //{
                        //    excelWorkBook.Close(false, Type.Missing, Type.Missing);
                        //    excelApp.Quit();
                        //    return;
                        //}
                        if (excelSheet.Range["B16"].Value.ToString().Contains("КП"))
                        {
                            int countNameInInvoice = 27;
                            do
                            {
                                //excelSheet = excelWorkBook.Worksheets[nameList];
                                ProductItem productItem = new ProductItem()
                                {
                                    Customer = null,
                                    Acct = regexAcct.Match(excelSheet.Range["B16"].Value).Value,
                                    Date = Convert.ToDateTime(regexDate.Match(excelSheet.Range["B16"].Value).Value),
                                    PartNumber = excelSheet.Range[$"AQ{countNameInInvoice}"].Value,
                                    Quanity = Convert.ToInt32(excelSheet.Range[$"AR{countNameInInvoice}"].Value),
                                    Price = Convert.ToDouble(excelSheet.Range[$"AS{countNameInInvoice}"].Value),
                                };
                                SaveData(productItem, ref sum, ref customer, ref acct, ref date, ref type);
                                items.Add(productItem);
                                //db.ProductItems.Add(productItem);
                                countNameInInvoice++;
                            } while ((excelSheet.Range[$"AR{countNameInInvoice}"].Value) != null);
                            countNameInInvoice = 27;
                        }
                        else if (excelSheet.Range["B5"].Value != null)
                        {
                            customerInfo = CustomerProcessor(excelSheet.Range[$"G11"].Value);

                            int countNameInInvoice = 16;
                            do
                            {
                                ProductItem productItem = new ProductItem()
                                {
                                    Acct = regexAcctInvoice.Match(excelSheet.Range["B5"].Value).Value,
                                    //Acct = excelSheet.Range["B5"].Value.Substring(17, 11),
                                    Date = Convert.ToDateTime(regexDate.Match(excelSheet.Range["B5"].Value).Value),
                                    Customer = excelSheet.Range[$"G11"].Value,
                                    PartNumber = excelSheet.Range[$"AQ{countNameInInvoice}"].Value,
                                    Quanity = Convert.ToInt32(excelSheet.Range[$"AR{countNameInInvoice}"].Value),
                                    Price = Convert.ToDouble(excelSheet.Range[$"AS{countNameInInvoice}"].Value),
                                    Type = nameList,

                                };
                                SaveData(productItem, ref sum, ref customer, ref acct, ref date, ref type);

                                items.Add(productItem);
                                //db.ProductItems.Add(productItem);
                                countNameInInvoice++;
                            } while ((excelSheet.Range[$"AR{countNameInInvoice}"].Value) != null);
                            countNameInInvoice = 16;
                        }

                        else if (excelSheet.Range["AQ25"].Value == null
                            && excelSheet.Range["AQ27"].Value == null
                            && excelSheet.Range["AQ15"].Value == null)
                        {
                            return;
                        }
                        else
                        {
                            char x;
                            if (excelSheet.Range["G22"].Value == null)
                            {
                                x = 'F';
                            }
                            else
                            {
                                x = 'G';
                            }
                            customerInfo = CustomerProcessor(excelSheet.Range[$"{x}22"].Value);
                            //customerInfos.Add(customerInfo);
                            if (excelSheet.Range["AQ25"].Value == null)
                            {
                                int countNameInInvoice = 27;
                                do
                                {
                                    ProductItem productItem = new ProductItem()
                                    {
                                        Acct = regexAcctInvoice.Match(excelSheet.Range["B16"].Value).Value,
                                        //Acct = excelSheet.Range["B16"].Value.Substring(17, 11),
                                        Date = Convert.ToDateTime(regexDate.Match(excelSheet.Range["B16"].Value).Value),
                                        Customer = excelSheet.Range[$"{x}22"].Value,
                                        PartNumber = excelSheet.Range[$"AQ{countNameInInvoice}"].Value,
                                        Quanity = Convert.ToInt32(excelSheet.Range[$"AR{countNameInInvoice}"].Value),
                                        Price = Convert.ToDouble(excelSheet.Range[$"AS{countNameInInvoice}"].Value),
                                        Type = nameList,

                                    };
                                    SaveData(productItem, ref sum, ref customer, ref acct, ref date, ref type);
                                    items.Add(productItem);
                                    //db.ProductItems.Add(productItem);
                                    countNameInInvoice++;
                                } while ((excelSheet.Range[$"AR{countNameInInvoice}"].Value) != null);
                                countNameInInvoice = 27;
                            }
                            else
                            {
                                int countNameInInvoice = 25;
                                do
                                {
                                    ProductItem productItem = new ProductItem()
                                    {
                                        Acct = regexAcctInvoice.Match(excelSheet.Range["B16"].Value).Value,
                                        //Acct = excelSheet.Range["B16"].Value.Substring(17, 11),
                                        Date = Convert.ToDateTime(regexDate.Match(excelSheet.Range["B16"].Value).Value),
                                        Customer = excelSheet.Range[$"{x}22"].Value,
                                        PartNumber = excelSheet.Range[$"AQ{countNameInInvoice}"].Value,
                                        Quanity = Convert.ToInt32(excelSheet.Range[$"AR{countNameInInvoice}"].Value),
                                        Price = Convert.ToDouble(excelSheet.Range[$"AS{countNameInInvoice}"].Value),
                                        Type = nameList
                                    };
                                    SaveData(productItem, ref sum, ref customer, ref acct, ref date, ref type);

                                    items.Add(productItem);
                                    //db.ProductItems.Add(productItem);

                                    countNameInInvoice++;

                                } while ((excelSheet.Range[$"AR{countNameInInvoice}"].Value) != null);
                                countNameInInvoice = 25;
                            }
                        }
                    }

                    else if (nameList == "КП" || nameList == "КП НДС 0%")
                    {
                        do
                        {
                            //excelSheet = excelWorkBook.Worksheets[nameList];
                            ProductItem productItem = new ProductItem()
                            {
                                Customer = null,
                                Acct = excelSheet.Range["B6"].Value.Substring(7, 11),
                                Date = Convert.ToDateTime(excelSheet.Range["B7"].Value.Substring(3, 10)),
                                PartNumber = excelSheet.Range[$"AQ{countNameInKP}"].Value,
                                Quanity = Convert.ToInt32(excelSheet.Range[$"AR{countNameInKP}"].Value),
                                Price = Convert.ToDouble(excelSheet.Range[$"AS{countNameInKP}"].Value),
                                Type = nameList,
                            };
                            SaveData(productItem, ref sum, ref customer, ref acct, ref date, ref type);

                            items.Add(productItem);
                            //db.ProductItems.Add(productItem);
                            countNameInKP++;
                        } while ((excelSheet.Range[$"AR{countNameInKP}"].Value) != null);
                        countNameInKP = 15;
                    }


                    
                }
                catch (Exception)
                {
                    //MessageBox.Show($"В файле {fileName} возникла ошибка, проверьте правильность заполения", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    
                }
                finally
                {
                    Invoice invoice = new Invoice()
                    {
                        Acct = acct,
                        //Company = customer,
                        Date = date,
                        Sum = sum,
                        Type = type,
                        Customer = customerInfo,
                    };
                    foreach (var item in db.Customers)
                    {
                        if (customerInfo.CustomerFull == item.CustomerFull)
                        {
                            invoice.Customer = item;
                        }
                    }
                    bool flag = false;
                    foreach (var item in db.Invoices)
                    {
                        if (invoice.Acct == item.Acct && invoice.Date == item.Date && invoice.Sum == item.Sum)
                        {
                            flag = true;
                            //MessageBox.Show($"{invoice.Acct} уже существует в базе", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            break;
                        }
                        else
                        {
                            flag = false;
                        }
                    }
                    if (!flag)
                    {
                        db.Invoices.Add(invoice);
                        foreach (var productItem in items)
                        {
                            db.ProductItems.Add(productItem);
                        }
                        db.SaveChanges();
                        excelWorkBook.Close(false, Type.Missing, Type.Missing);
                        excelApp.Quit();
                    }
                    else
                    {
                        excelWorkBook.Close(false, Type.Missing, Type.Missing);
                        excelApp.Quit();
                    }
                }
            }
        }

        public Customer CustomerProcessor(string fullCustomerInfo)
        {

            String customerRawData = fullCustomerInfo;

            string companyName;
            string inn;
            string adress;
            string tel;
            string customerFull;
            customerFull = customerRawData;

            Regex regexCompanyName = new Regex(RegularFormular.CompanyNamePattern);
            companyName = regexCompanyName.Match(customerRawData).Value;
            customerRawData = regexCompanyName.Replace(customerRawData, string.Empty, 1);

            Regex regexInnOrBin = new Regex(RegularFormular.InnOrBinnPattern);
            inn = regexInnOrBin.Match(customerRawData).Value;
            customerRawData = regexInnOrBin.Replace(customerRawData, string.Empty, 1);
            Regex regexclearRnn = new Regex(RegularFormular.ClearRnnPattern);
            customerRawData = regexclearRnn.Replace(customerRawData, string.Empty, 1);

            Regex regexCleanKpp = new Regex(RegularFormular.ClearKppPattern);
            customerRawData = regexCleanKpp.Replace(customerRawData, string.Empty, 1);

            Regex regexAdress = new Regex(RegularFormular.IndexPattern);
            adress = regexAdress.Match(customerRawData).Value;
            customerRawData = regexAdress.Replace(customerRawData, string.Empty, 1);

            Regex regexTel = new Regex(RegularFormular.TelephonPattern);
            tel = regexTel.Match(customerRawData).Value;
            customerRawData = regexTel.Replace(customerRawData, string.Empty, 1);
            customerRawData = regexTel.Replace(customerRawData, string.Empty, 1);

            Regex regexClear = new Regex(RegularFormular.Clear);
            customerRawData = regexClear.Replace(customerRawData, string.Empty, 1);

            Regex regexClear2 = new Regex(RegularFormular.Clear2);
            customerRawData = regexClear2.Replace(customerRawData, string.Empty, 1);
            adress += customerRawData;
            tel = tel.TrimStart('.');
            tel = tel.TrimStart(':');
            tel = tel.Trim();
            adress = adress.TrimStart(',');
            adress = adress.Trim();
            Customer customerInfo = new Customer()
            {
                Adress = adress,
                CompanyName = companyName,
                CustomerFull = customerFull,
                Inn = inn,
                Tel = tel
            };
            //Попытаться разделить метод на несколько
            return customerInfo;
        }
    }
}

