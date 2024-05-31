using System;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;


namespace FileCreators
{
    public class WordCreator
    {
        public static void Export(string path)
        {
            Word.Application wdApp = new Word.Application();
            Word.Document doc = wdApp.Documents.Add();
            Word.Paragraph p = doc.Paragraphs.Add();

            p.Range.Text = "Text text text";
            p.Range.InsertParagraphAfter();

            Word.Table table = doc.Tables.Add(p.Range, 3, 3);

            table.Cell(1, 1).Range.Text = "Заголовок";
            table.Cell(1, 2).Range.Text = "Заголовок";

            table.Cell(2, 1).Range.Text = "adadad";
            table.Cell(2, 2).Range.Text = "fafasd";

            table.Borders.Enable = 1;
            table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            try
            {
                doc.SaveAs(path);
                Console.WriteLine("Word документ створений");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Виникла помилка при збереженнi документу Word за шляхом {path}");
            }
            finally
            {
                doc.Close();
                wdApp.Quit();
                Marshal.ReleaseComObject(doc);
                Marshal.ReleaseComObject(wdApp);
                Marshal.ReleaseComObject(p);
                Marshal.ReleaseComObject(table);
            }
        }
    }
}
