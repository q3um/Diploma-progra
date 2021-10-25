using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace parserExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            List<ProductItem> items = new List<ProductItem>();
            List<CustomerInfo> customerInfos = new List<CustomerInfo>();
            Parser parser = new Parser();
            //string fileName = @"c:\Users\asus\Desktop\Parsing\72003101787.xlsm";
            string[] fileList = Directory.GetFiles(@"C:\Users\asus\Desktop\Parsing\ToParseCSV", "*.xlsm", SearchOption.AllDirectories);
            foreach (string fileName in fileList)
            {
                if (fileName.Contains("~$"))
                {
                    continue;
                }
                parser.ProductItemProcessor(fileName, items, customerInfos);
            }
            parser.PrintProductItems(items);
            parser.PrintProductCustomerInfos(customerInfos);
            //parser.CustomerProcessorToTxt();

            //Contract.formFillingContract();
            Console.WriteLine("Complete");
            Console.ReadKey();
        }
    }
}
