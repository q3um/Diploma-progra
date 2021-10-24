using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace parserExcel
{
    class Parser
    {

        public void ProductItemProcessor(string fileName, List<ProductItem> items, List<CustomerInfo> customerInfos)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkBook;
            Excel.Worksheet excelSheet;
            int countNameInKP = 15; // строка с которой начинаются товары в кп

            excelWorkBook = excelApp.Workbooks.Open(fileName);
            string nameList = (excelApp.Sheets[1] as Excel.Worksheet).Name;
            excelSheet = excelWorkBook.Worksheets[nameList];

            CustomerInfo customerInfo = new CustomerInfo();
            if (nameList == "НДС внутри" || nameList == "НДС сверху" || nameList == "НДС 0%")
            {
                if (excelSheet.Range["B16"].Value.Contains("КП"))
                {
                    int countNameInInvoice = 27;
                    do
                    {
                        RegularFormular regularFormular = new RegularFormular();
                        //excelSheet = excelWorkBook.Worksheets[nameList];
                        ProductItem productItem = new ProductItem();
                        productItem.Customer = null;
                        Regex regexAcct = new Regex(regularFormular.AcctInKpInvoice);
                        productItem.Acct = Convert.ToInt64(regexAcct.Match(excelSheet.Range["B16"].Value).Value);
                        Regex regexDate = new Regex(regularFormular.DateInKPSheet);
                        productItem.Date = regexDate.Match(excelSheet.Range["B16"].Value).Value;
                        productItem.PartNumber = excelSheet.Range[$"AQ{countNameInInvoice}"].Value;
                        productItem.Quanity = Convert.ToInt32(excelSheet.Range[$"AR{countNameInInvoice}"].Value);
                        productItem.Price = Convert.ToDouble(excelSheet.Range[$"AS{countNameInInvoice}"].Value);
                        productItem.Sum = productItem.Price * productItem.Quanity;
                        productItem.Type = "КП";
                        items.Add(productItem);
                        countNameInInvoice++;
                    } while ((excelSheet.Range[$"AR{countNameInInvoice}"].Value) != null);
                    countNameInInvoice = 27;
                }

                else if (excelSheet.Range["AQ25"].Value == null & excelSheet.Range["AQ27"].Value == null & excelSheet.Range["AQ15"].Value == null)
                {
                    return;
                }
                else
                {
                    customerInfo = CustomerProcessor(excelSheet.Range["G22"].Value);
                    customerInfos.Add(customerInfo);
                    if (excelSheet.Range["AQ25"].Value == null)
                    {
                        int countNameInInvoice = 27; //строка с которой начинаются товары в счете
                        do
                        {
                            ProductItem productItem = new ProductItem();
                            productItem.Acct = Convert.ToInt64(excelSheet.Range["B16"].Value.Substring(17, 11));
                            productItem.Date = excelSheet.Range["B16"].Value.Substring(32, 10);
                            productItem.Customer = excelSheet.Range["G22"].Value;
                            productItem.PartNumber = excelSheet.Range[$"AQ{countNameInInvoice}"].Value;
                            productItem.Quanity = Convert.ToInt32(excelSheet.Range[$"AR{countNameInInvoice}"].Value);
                            productItem.Price = Convert.ToDouble(excelSheet.Range[$"AS{countNameInInvoice}"].Value);
                            productItem.Sum = productItem.Price * productItem.Quanity;
                            productItem.Type = nameList;
                            items.Add(productItem);
                            countNameInInvoice++;
                        } while ((excelSheet.Range[$"AR{countNameInInvoice}"].Value) != null);
                        countNameInInvoice = 27;
                    }
                    else
                    {
                        int countNameInInvoice = 25; //строка с которой начинаются товары в счете
                        do
                        {
                            ProductItem productItem = new ProductItem();
                            productItem.Acct = Convert.ToInt64(excelSheet.Range["B16"].Value.Substring(17, 11));
                            productItem.Date = excelSheet.Range["B16"].Value.Substring(32, 10);
                            productItem.Customer = excelSheet.Range["G22"].Value;
                            productItem.PartNumber = excelSheet.Range[$"AQ{countNameInInvoice}"].Value;
                            productItem.Quanity = Convert.ToInt32(excelSheet.Range[$"AR{countNameInInvoice}"].Value);
                            productItem.Price = Convert.ToDouble(excelSheet.Range[$"AS{countNameInInvoice}"].Value);
                            productItem.Sum = productItem.Price * productItem.Quanity;
                            productItem.Type = nameList;
                            items.Add(productItem);
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
                    ProductItem productItem = new ProductItem();
                    productItem.Customer = null;
                    productItem.Acct = Convert.ToInt64(excelSheet.Range["B6"].Value.Substring(7, 11));
                    productItem.Date = (excelSheet.Range["B7"].Value.Substring(3, 10));
                    productItem.PartNumber = excelSheet.Range[$"AQ{countNameInKP}"].Value;
                    productItem.Quanity = Convert.ToInt32(excelSheet.Range[$"AR{countNameInKP}"].Value);
                    productItem.Price = Convert.ToDouble(excelSheet.Range[$"AS{countNameInKP}"].Value);
                    productItem.Sum = productItem.Price * productItem.Quanity;
                    productItem.Type = nameList;
                    items.Add(productItem);
                    countNameInKP++;
                } while ((excelSheet.Range[$"AR{countNameInKP}"].Value) != null);
                countNameInKP = 15;
            }
            excelWorkBook.Close(false, Type.Missing, Type.Missing);
            excelApp.Quit();
        }

        public CustomerInfo CustomerProcessor(string fullCustomerInfo)
        {

            String customer = fullCustomerInfo;
            CustomerInfo customerInfo = new CustomerInfo();

            customerInfo.Customer = customer;
            RegularFormular regularFormular = new RegularFormular();

            Regex regexCompanyName = new Regex(regularFormular.CompanyNamePattern);
            customerInfo.CompanyName = regexCompanyName.Match(customer).Value;
            customer = regexCompanyName.Replace(customer, string.Empty, 1);

            Regex regexInnOrBin = new Regex(regularFormular.InnOrBinnPattern);
            customerInfo.Inn = regexInnOrBin.Match(customer).Value;
            customer = regexInnOrBin.Replace(customer, string.Empty, 1);
            Regex regexclearRnn = new Regex(regularFormular.ClearRnnPattern);
            customer = regexclearRnn.Replace(customer, string.Empty, 1);

            Regex regexCleanKpp = new Regex(regularFormular.ClearKppPattern);
            customer = regexCleanKpp.Replace(customer, string.Empty, 1);

            Regex regexAdress = new Regex(regularFormular.IndexPattern);
            customerInfo.Adress = regexAdress.Match(customer).Value;
            customer = regexAdress.Replace(customer, string.Empty, 1);

            Regex regexTel = new Regex(regularFormular.TelephonPattern);
            customerInfo.Tel = regexTel.Match(customer).Value;
            customer = regexTel.Replace(customer, string.Empty, 1);
            customer = regexTel.Replace(customer, string.Empty, 1);

            Regex regexClear = new Regex(@"(/)?(,)?\s{0,2}(,)?\s{0,2}((тел)|(Тел)|(Доб)|(доб)|(Моб)|(моб)|(факс)|(факс))(:)?(.)?\s{0,2}\d{0,4}(-)?\d{0,4}(,)?");
            customer = regexClear.Replace(customer, string.Empty, 1);

            Regex regexClear2 = new Regex(@"($,){0,2}");
            customer = regexClear2.Replace(customer, string.Empty, 1);
            customerInfo.Adress += customer;
            return customerInfo;
        }

        public void ReadCustomerInSheets(string fileName, ref string customer)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkBook;
            Excel.Worksheet excelSheet;

            excelWorkBook = excelApp.Workbooks.Open(fileName);
            string nameList = (excelApp.Sheets[1] as Excel.Worksheet).Name;

            if (nameList == "НДС внутри" || nameList == "НДС сверху" || nameList == "НДС 0%")
            {
                excelSheet = excelWorkBook.Worksheets[nameList];
                customer = (excelSheet.Range["G22"].Value) ?? string.Empty;
            }
            else
            {
                excelApp.Quit();
            }
            excelApp.Quit();
        }

        public void CustomerProcessorToTxt()
        {
            //List<string> CSV_String = new List<string>();
            List<CustomerInfo> customerInfoList = new List<CustomerInfo>();
            string[] fileList = Directory.GetFiles(@"C:\Users\asus\Desktop\Parsing\ToParseCSV", "*.xlsm", SearchOption.AllDirectories);
            String customer;
            foreach (string fileToRead in fileList)
            {
                if (fileToRead.Contains("~$"))
                {
                    continue;
                }
                CustomerInfo customerInfo = new CustomerInfo();

                customer = String.Empty;
                ReadCustomerInSheets(fileToRead, ref customer);
                //CSV_String.Add(customer);
                customerInfo.Customer = customer;
                RegularFormular regularFormular = new RegularFormular();

                Regex regexCompanyName = new Regex(regularFormular.CompanyNamePattern);
                customerInfo.CompanyName = regexCompanyName.Match(customer).Value;
                customer = regexCompanyName.Replace(customer, string.Empty, 1);

                Regex regexInnOrBin = new Regex(regularFormular.InnOrBinnPattern);
                customerInfo.Inn = regexInnOrBin.Match(customer).Value;
                customer = regexInnOrBin.Replace(customer, string.Empty, 1);
                Regex regexclearRnn = new Regex(regularFormular.ClearRnnPattern);
                customer = regexclearRnn.Replace(customer, string.Empty, 1);



                Regex regexCleanKpp = new Regex(regularFormular.ClearKppPattern);
                customer = regexCleanKpp.Replace(customer, string.Empty, 1);

                Regex regexAdress = new Regex(regularFormular.IndexPattern);
                customerInfo.Adress = regexAdress.Match(customer).Value;
                customer = regexAdress.Replace(customer, string.Empty, 1);

                Regex regexTel = new Regex(regularFormular.TelephonPattern);
                customerInfo.Tel = regexTel.Match(customer).Value;
                customer = regexTel.Replace(customer, string.Empty, 1);
                customer = regexTel.Replace(customer, string.Empty, 1);

                Regex regexClear = new Regex(@"(/)?(,)?\s{0,2}(,)?\s{0,2}((тел)|(Тел)|(Доб)|(доб)|(Моб)|(моб)|(факс)|(факс))(:)?(.)?\s{0,2}\d{0,4}(-)?\d{0,4}(,)?");
                customer = regexClear.Replace(customer, string.Empty, 1);

                Regex regexClear2 = new Regex(@"($,){0,2}");
                customer = regexClear2.Replace(customer, string.Empty, 1);
                customerInfo.Adress += customer;
                customerInfoList.Add(customerInfo);
            }
            StreamWriter streamWriter = new StreamWriter(@"C:\Users\asus\Desktop\Parsing\ToParseCSV\result.txt");
            foreach (var item in customerInfoList)
            {
                streamWriter.WriteLine($"Изначальная строка: {item.Customer}");
                streamWriter.WriteLine($"Название компании: {item.CompanyName}");
                streamWriter.WriteLine($"ИНН: {item.Inn}");
                streamWriter.WriteLine($"Адресс: {item.Adress}");
                streamWriter.WriteLine($"Телефон: {item.Tel}");
                streamWriter.WriteLine("========================================");
            }
            ////for (int i = 0; i < CSV_String.Count; i++)
            ////{
            ////    streamWriter.WriteLine(CSV_String[i]);
            ////}
            streamWriter.Close();
            Console.WriteLine("Complete");
        }

        public void productProcessor(string fileName)
        {
            int countNameInKP = 15;
            int countNameInInvoice = 27;
            int countNameInInvoiceOld = 25;
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkBook;
            Excel.Worksheet excelSheet;

            excelWorkBook = excelApp.Workbooks.Open(fileName);
            string nameList = (excelApp.Sheets[1] as Excel.Worksheet).Name;
            string fullProductInfo;
            excelSheet = excelWorkBook.Worksheets[nameList];
            if (nameList == "НДС внутри" || nameList == "НДС сверху" || nameList == "НДС 0%")
            {

            }
            else if (nameList == "КП" || nameList == "КП НДС 0%")
            {

            }
            else if (true)
            {

            }
            fullProductInfo = (excelSheet.Range[$"AQ{countNameInInvoice}"].Value) ?? string.Empty;

            excelApp.Quit();

        }

        public void printProductItems(List<ProductItem> items)
        {
            foreach (var item in items)
            {
                Console.WriteLine($"{item.Type} {item.Acct} {item.Date} {item.Customer} {item.PartNumber} {item.Quanity} {item.Price} {item.Sum}");
            }
        }
        public void printProductCustomerInfos(List<CustomerInfo> customerInfos)
        {
            foreach (var item in customerInfos)
            {
                Console.WriteLine($"{item.Customer}\n{item.CompanyName}\n{item.Adress}\n{item.Inn}\n{item.Tel}");
            }
        }
    }
}
