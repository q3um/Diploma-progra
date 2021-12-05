using System;
using System.Collections.Generic;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace ContractFill
{
    class wordHelper
    {
        private FileInfo _fileInfo;

        public wordHelper(string fileName)
        {
            if (File.Exists(fileName))
            {
                _fileInfo = new FileInfo(fileName);
            }
            else
            {
                throw new ArgumentException("File not found");
            }
        }

        public string FioSokr(string FullNameFIO)
        {
            string[] Fio = FullNameFIO.Split(' ');
            string FioSokr;
            if (Fio.Length == 2)
            {
            FioSokr = Fio[0] + ' ' + Fio[1][0] + '.';
            }
            else
            {
            FioSokr = Fio[0] + ' ' + Fio[1][0] + '.' + Fio[2][0] + '.';
            }

            return FioSokr;
        }

        internal bool Process(Dictionary<string, string> items, List<ProductItem> productItems)
        {
            Word.Application app = null;
            try
            {
                app = new Word.Application();
                Object file = _fileInfo.FullName;
                Object missing = Type.Missing;

                app.Documents.Open(file);

                foreach (var item in items)
                {
                    Word.Find find = app.Selection.Find;
                    find.Text = item.Key;
                    find.Replacement.Text = item.Value;
                    object wrap = Word.WdFindWrap.wdFindContinue;
                    object replace = Word.WdReplace.wdReplaceAll;

                    find.Execute(FindText: Type.Missing,
                       MatchCase: false,
                       MatchWholeWord: false,
                       MatchWildcards: false,
                       MatchSoundsLike: missing,
                       MatchAllWordForms: false,
                       Forward: true,
                       Wrap: wrap,
                       Format: false,
                       ReplaceWith: missing, Replace: replace);
                }

                Word.Document document = app.Documents.Add(file);

                Word.Table table = document.Tables[2];
                int countProduct = 1;
                int countRows = 2;
                foreach (var item in productItems)
                {
                    table.Rows.Add(document.Tables[2].Rows[countRows]);
                    table.Rows[countRows].Cells[1].Range.Text = Convert.ToString(countProduct);
                    table.Rows[countRows].Cells[2].Range.Text = Convert.ToString(item.Partnumber);
                    table.Rows[countRows].Cells[3].Range.Text = "шт.";
                    table.Rows[countRows].Cells[4].Range.Text = Convert.ToString(item.Quanity);
                    table.Rows[countRows].Cells[5].Range.Text = item.Price.ToString("F" + 2);
                    table.Rows[countRows].Cells[6].Range.Text = item.Sum.ToString("F" + 2);
                    countRows++;
                    countProduct++;
                }
                table.Rows[countRows].Delete();



                //document.Close();
                Object newFileName = Path.Combine(_fileInfo.DirectoryName, DateTime.Now.ToString("yyyyMMdd HHmmss ") + _fileInfo.Name);
                app.ActiveDocument.SaveAs2(newFileName);
                app.ActiveDocument.Close();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                if (app != null)
                {
                    app.Quit(false);

                }
            }

            return false;
        }
    }
}
