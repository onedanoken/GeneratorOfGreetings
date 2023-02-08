using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Reflection.Emit;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;


namespace lab1
{
    class Program
    {
        static void Main(string[] args)
        {
            List<(string, int)> list = null;
            list = LoadFromExcel(@"C:\input.xlsx");
        }

        public static List<(string, int)> LoadFromExcel(string fileName)
        {
            Excel.Application excel = null; 
            Excel.Workbook wbk = null;

            try
            {
                excel = new Excel.Application();
                wbk = excel.Workbooks.Open(fileName);
            }
            catch (Exception e) 
            {
                Console.WriteLine(e.Message);            
            }


            List<(string, int)> list = new List<(string, int)>();
            if (excel != null && wbk != null)
            {
                Excel.Worksheet sheetNames = (Excel.Worksheet)wbk.Sheets[1]; // Получить доступ к именам
                //Excel.Worksheet sheetWishes = (Excel.Worksheet)wbk.Sheets["2"]; // Получить доступ к поздравлениям
                //Excel.Worksheet Conf = (Excel.Worksheet)wbk.Sheets["3"]; // Получить доступ к конфигу

                int rows = sheetNames.Rows.Count;
                int cols = sheetNames.Columns.Count;

                for (int i = 0; i < cols; i++) 
                {
                    var a = ((string)sheetNames.Cells[i, 0], (int)sheetNames.Cells[i, 1]);
                    list.Add(a);
                }

            }
            if (excel != null)
                excel.Quit();

            return list;
        }
        
        public static void ExportToWord(List<string> list)
        {
            Word.Application word = null;
            Word.Document doc = null;
            object oEndOfDoc = "\\endofdoc";

            try
            {
                word = new Word.Application();
                word.Visible = true;
                doc = word.Documents.Add();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
