using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WindowsFormsApp1
{
    class Contract
    {
        private static void ReplaceStub(string stubToReplace, string text, Word.Document wordDocument)
        {
            var range = wordDocument.Content;
            range.Find.ClearFormatting();
            object wdReplaceAll = Word.WdReplace.wdReplaceAll;
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text, Replace: wdReplaceAll);
        }
        public static void formFillingContract()
        {
            Word.Application wordApp = new Word.Application();
            string textBox_fio = "Иванов Иван Иваныч";
            try
            {
                var wordDoc = wordApp.Documents.Open(@"c:\Users\asus\Desktop\Parsing\Договор.docx");
                ReplaceStub("{fio}", textBox_fio, wordDoc);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            wordApp.Quit();
        }
    }
}
