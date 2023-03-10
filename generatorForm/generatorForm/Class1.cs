using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Metrics;
using System.Net;
using System.Reflection.Emit;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;


namespace generatorForm
{
    // Класс данных для людей
    class People
    {
        private string name;
        private int id;

        public string Name { get { return name; } set { this.name = value; } }
        public int Id { get { return id; } set { this.id = value; } }

        public People(string name, int id)
        {
            this.name = name;
            this.id = id;
        }
    }

    // Класс данных для хранения триад поздравлений
    class Triads
    {
        private string firstGreeting;
        private string secondGreeting;
        private string thirdGreeting;

        public string FirstGreeting { get { return this.firstGreeting; } set { this.firstGreeting = value; } }
        public string SecondGreeting { get { return this.secondGreeting; } set { this.secondGreeting = value; } }
        public string ThirdGreeting { get { return this.thirdGreeting; } set { this.thirdGreeting = value; } }

        public Triads(string firstGreeting, string secondGreeting, string thirdGreeting)
        {
            this.firstGreeting = firstGreeting;
            this.secondGreeting = secondGreeting;
            this.thirdGreeting = thirdGreeting;
        }
    }

    class WorkWithFiles
    {
        public static int Start()
        {
            int code = 0; // по умолчанию ноль - всё хорошо;
            string fileName = @"C:\СЯиТП\лаба 1\input.xlsx"; // путь к конфигу
            List<People> names = null;
            List<List<string>> greetings = null;
            List<Triads> triads = null;
            names = LoadNamesFromExcel(fileName);
            if (names == null)
            {
                code = 1; // 1 - проблема с открытием файла excel;
                return code;
            }
            greetings = LoadGreetingsFromExcel(fileName);
            if (greetings == null)
            {
                code = 1; 
                return code;
            }
            int n = names.Count;
            triads = CreateTriads(greetings, n);
            if (triads == null)
            {
                code = 2; // 2 - проблема с создание триад - нехватка поздравлений;
                return code;
            }
            code = ExportToWord(names, triads, fileName);
            return code;
        }

        public static List<People> LoadNamesFromExcel(string fileName)
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
                return null;
            }

            Excel.Worksheet sheetNames = wbk.Worksheets[1]; // Получить доступ к именам
            var lastRow = sheetNames.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
            List<People> names = new List<People>();

            for (int i = 1; i <= lastRow; i++)
            {
                string name = (sheetNames.Cells[i, 1]).Value2;
                int id = Convert.ToInt32((sheetNames.Cells[i, 2]).Value2);
                People person = new People(name, id);
                names.Add(person);
            }

            if (excel != null)
            {
                excel.Quit();
            }

            Process[] prc;
            prc = Process.GetProcessesByName("EXCEL");
            foreach (Process proc in prc)
            {
                proc.Kill();
            }

            return names;
        }

        public static List<List<string>> LoadGreetingsFromExcel(string fileName)
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
                return null;
            }

            Excel.Worksheet sheetNames = wbk.Worksheets[2]; // получить доступ к поздравлениям
            var lastRow = sheetNames.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
            var lastColumn = sheetNames.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Column;
            List<List<string>> greetings = new List<List<string>>();
            List<string> tmp = null;

            for (int i = 1; i <= lastColumn; i++)
            {
                tmp = new List<string>();
                for (int j = 2; j <= lastRow; j++)
                {
                    var a = (sheetNames.Cells[j, i]).Value2;
                    if (a == null)
                    {
                        break;
                    }
                    tmp.Add(a);
                }
                greetings.Add(tmp);
            }

            if (excel != null)
            {
                excel.Quit();
            }

            Process[] prc;
            prc = Process.GetProcessesByName("EXCEL");
            foreach (Process proc in prc)
            {
                proc.Kill();
            }

            return greetings;

        }

        public static List<Triads> CreateTriads(List<List<string>> greetings, int n)
        {
            // Проверка количества составить триады.
            int maxCountTriad = 0;
            for (int i = 0; i < greetings.Count - 2; ++i)
            {
                for (int j = i + 1; j < greetings.Count - 1; ++j)
                {
                    for (int k = j + 1; k < greetings.Count; ++k)
                    {
                        maxCountTriad += 
                            greetings[i].Count *
                            greetings[j].Count *
                            greetings[k].Count;
                    }
                }
            }
            if (maxCountTriad < n)
            {
                return null; // Если триад не хватает - уведомляем пользователя об этом.
            }

            // Словарь для использованных поздравлений. False - поздравление не использовано
            List<Dictionary<string, bool>> usedGreetings = new List<Dictionary<string, bool>>();
            foreach (var categories in greetings)
            {
                Dictionary<string, bool> tmp= new Dictionary<string, bool>();
                foreach (var g in categories)
                {
                    tmp.Add((string)g, false);
                }
                usedGreetings.Add(tmp);
            }

            Random rnd = new Random();
            List<Triads> triads = new List<Triads>();
            int gCount = greetings.Count - 1;
            while (triads.Count < n)
            {
                int index1 = rnd.Next(gCount);
                int index2 = rnd.Next(gCount);
                int index3 = rnd.Next(gCount);
                if (index1 == index2 || index1 == index3 || index2 == index3)
                {
                    continue;
                }
                var lst1 = greetings[index1];
                var lst2 = greetings[index2];
                var lst3 = greetings[index3];
                int j1 = rnd.Next(lst1.Count - 1);
                int j2 = rnd.Next(lst2.Count - 1);
                int j3 = rnd.Next(lst3.Count - 1);
                if (usedGreetings[index1][lst1[j1]] == true || usedGreetings[index2][lst2[j2]] == true 
                    || usedGreetings[index3][lst3[j3]] == true) { continue; }

                Triads triada = new Triads(lst1[j1], lst2[j2], lst3[j3]);

                if (triads.Contains(triada))
                {
                    continue;
                }
                triads.Add(triada);

                // Проверка на то, что в категории использовались все поздравления.
                foreach (var category in usedGreetings)
                {
                    int count = 0; // кол-во использованных
                    foreach (var greeting in category)
                    {
                        if (greeting.Value == true) { count += 1; }
                    }
                    if (count == category.Count) // Если все использованы - меняем на false
                    {
                        foreach (var key in category)  { category[key.Key] = false; }
                    }
                }
            }
            return triads;
        }


        public static int ExportToWord(List<People> names, List<Triads> triads, string fileNameExcel)
        {
            try
            {
                Word.Application word = null;
                Word.Document originalDoc = null;
                Word.Document newDoc = null;

                //Console.WriteLine("Идёт генерация...");
                //Console.WriteLine("Пожалуйста, подождите.");

                // Работа с Excel
                Excel.Application excel = null;
                Excel.Workbook wbk = null;

                excel = new Excel.Application();
                wbk = excel.Workbooks.Open(fileNameExcel);
               
                Excel.Worksheet sheetNames = wbk.Worksheets[3]; // Получить доступ к конфигу

                // Получаем доступ к шаблону и выходной директории и шрифту
                string originalFile = (sheetNames.Cells[4, 2]).Value2; // путь к шаблону
                string path = (sheetNames.Cells[5, 2]).Value2; // путь к папке out
                var fontName = (sheetNames.Cells[3, 2]).Value2; // путь к шрифту

                // Проверка наличия выходной директории
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);

                word = new Word.Application();

                // Создаём выходной файл 
                string newFileName = "result_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".docx";
                string newFile = path + "\\" + newFileName;
                originalDoc = word.Documents.Open(originalFile, Visible: false);
                newDoc = word.Documents.Add();
                int count = names.Count;

                var first_bookmark_range = newDoc.Words.Last;
                first_bookmark_range.Font.Name = fontName;

                // Генератор поздравлений
                for (int i = 0; i < count; i++)
                {
                    originalDoc.Content.FormattedText.Copy();
                    newDoc.Words.Last.Paste();
                    foreach (Word.Bookmark bookmark in newDoc.Bookmarks)
                    {
                        string bookmarkName = bookmark.Name;
                        bookmark.Range.Font.Name = fontName;
                        bookmark.Range.Font.Size = 14;
                        switch (bookmarkName)
                        {
                            case "person_name":
                                if (names[i].Id == 1)
                                {
                                    var treatmeant = (sheetNames.Cells[2, 2]).Value2;
                                    bookmark.Range.Text = treatmeant + " " + names[i].Name + "!";
                                }
                                else
                                {
                                    var treatmeant = (sheetNames.Cells[2, 3]).Value2;
                                    bookmark.Range.Text = treatmeant + " " + names[i].Name + "!";
                                }
                                break;
                            case "standart_greeting":
                                var Event = (sheetNames.Cells[1, 2]).Value2;
                                bookmark.Range.Text = "Поздравляю вас с " + Event + "!";
                                break;
                            case "first_greeting":
                                bookmark.Range.Text = "Желаю " + triads[i].FirstGreeting + ",";
                                break;
                            case "second_greeting":
                                bookmark.Range.Text = triads[i].SecondGreeting + ",";
                                break;
                            default:
                                bookmark.Range.Text = triads[i].ThirdGreeting;
                                break;
                        }
                    }
                    if (i != count - 1)
                    {
                        newDoc.Words.Last.InsertBreak(WdBreakType.wdPageBreak);
                    }
                }
                newDoc.SaveAs2(newFile);
                newDoc.Close();

                // Подчищаем данные

                if (word != null)
                {
                    word.Quit();
                }

                if (excel != null)
                {
                    excel.Quit();
                }

                Process[] prc;
                prc = Process.GetProcessesByName("EXCEL");
                foreach (Process proc in prc)
                {
                    proc.Kill();
                }

                //Console.WriteLine("Генерация завершена успешно!");
                return 0;
            }
            catch(Exception ex)
            {
                return 3; // 3 - Проблема с загрузкой в word
            }
        }
    }
}