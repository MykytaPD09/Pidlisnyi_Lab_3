using Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.InteropServices;

namespace FileCreators
{
    public class ExelCreator
    {
        public static void Export(string path)
        {
            Application xlApp = new Application();
            Workbook xlBook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet xlSheet = (Worksheet)xlBook.Worksheets[1];

            xlSheet.Cells[1, 1] = "ddddd";
            xlSheet.Cells[1, 2] = "dadadad";

            try
            {
                xlBook.SaveAs(path);
                Console.WriteLine("Exel документ створений");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Виникла помилка при збереженнi Exel документу за шляхом {path}");
            }
            finally
            {
                xlBook.Close(false);
                xlApp.Quit();
                Marshal.ReleaseComObject(xlSheet);
                Marshal.ReleaseComObject(xlBook);
                Marshal.ReleaseComObject(xlApp);
            }
        }
    }
}
