using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp1
{
    class ParserInvoice
    {
        private FileInfo _fileInfo;

        public ParserInvoice(string file)
        {
            if (File.Exists(file))
            {
                _fileInfo = new FileInfo(file);
            }
            else
            {
                throw new ArgumentException("File not found");
            }
        }
        public void ProductItemProcessor(List<ProductItem> items)
        {
            string file = _fileInfo.FullName;
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkBook;
            Excel.Worksheet excelSheet;
            int countNameInKP = 15; // строка с которой начинаются товары в кп

            excelWorkBook = excelApp.Workbooks.Open(file);
            string nameList = (excelApp.Sheets[1] as Excel.Worksheet).Name;
            excelSheet = excelWorkBook.Worksheets[nameList];

            if (nameList == "НДС внутри" || nameList == "НДС сверху" || nameList == "НДС 0%")
            {
                if (excelSheet.Range["B16"].Value.Contains("КП"))
                {
                    int countNameInInvoice = 27;
                    do
                    {
                        ProductItem productItem = new ProductItem()
                        {
                            PartNumber = excelSheet.Range[$"AQ{countNameInInvoice}"].Value,
                            Quanity = Convert.ToInt32(excelSheet.Range[$"AR{countNameInInvoice}"].Value),
                            Price = Convert.ToDouble(excelSheet.Range[$"AS{countNameInInvoice}"].Value),
                        };
                        productItem.Sum = productItem.Price * productItem.Quanity;
                        items.Add(productItem);
                        countNameInInvoice++;
                    } while ((excelSheet.Range[$"AR{countNameInInvoice}"].Value) != null);
                    countNameInInvoice = 27;
                }

                else if (excelSheet.Range["AQ25"].Value == null
                    && excelSheet.Range["AQ27"].Value == null
                    && excelSheet.Range["AQ15"].Value == null)
                {
                    return;
                }
                else
                {
                    if (excelSheet.Range["AQ25"].Value == null)
                    {
                        int countNameInInvoice = 27;
                        do
                        {
                            ProductItem productItem = new ProductItem()
                            {
                                PartNumber = excelSheet.Range[$"AQ{countNameInInvoice}"].Value,
                                Quanity = Convert.ToInt32(excelSheet.Range[$"AR{countNameInInvoice}"].Value),
                                Price = Convert.ToDouble(excelSheet.Range[$"AS{countNameInInvoice}"].Value),
                            };
                            productItem.Sum = productItem.Price * productItem.Quanity;

                            items.Add(productItem);
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
                                PartNumber = excelSheet.Range[$"AQ{countNameInInvoice}"].Value,
                                Quanity = Convert.ToInt32(excelSheet.Range[$"AR{countNameInInvoice}"].Value),
                                Price = Convert.ToDouble(excelSheet.Range[$"AS{countNameInInvoice}"].Value)
                            };
                            productItem.Sum = productItem.Price * productItem.Quanity;
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
                    ProductItem productItem = new ProductItem()
                    {
                        PartNumber = excelSheet.Range[$"AQ{countNameInKP}"].Value,
                        Quanity = Convert.ToInt32(excelSheet.Range[$"AR{countNameInKP}"].Value),
                        Price = Convert.ToDouble(excelSheet.Range[$"AS{countNameInKP}"].Value)
                    };
                    productItem.Sum = productItem.Price * productItem.Quanity;
                    items.Add(productItem);
                    countNameInKP++;
                } while ((excelSheet.Range[$"AR{countNameInKP}"].Value) != null);
                countNameInKP = 15;
            }
            excelWorkBook.Close(false, Type.Missing, Type.Missing);
            excelApp.Quit();
        }
    }
}
