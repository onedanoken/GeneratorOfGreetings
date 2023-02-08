using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficeTest
{
	class Program
	{
		public static void Main(string[] args)
		{
			Console.WriteLine("Hello World!");
			
			List<string> list = null;
			
			/*
			list = new List<string>();
			
			list.Add("Раз");
			list.Add("Два");
			list.Add("Три");
			*/
			
			list = LoadFromExcel(@"C:\1.xlsx");
			
			ExportToWord(list);
			
			Console.ReadKey(true);
		}
		
		public static List<string> LoadFromExcel(string filename)
		{
			Excel.Application excel = null;
			Excel.Workbook wbk = null;
			
			try {
				excel = new Excel.Application();
				excel.Visible = true;
				wbk = excel.Workbooks.Open(filename);
			} catch (Exception ex) {
				Console.WriteLine(ex);
			}
			
			List<string> list = new List<string>();
			
			if (excel != null && wbk != null)
			{
				Excel.Worksheet ws = (Excel.Worksheet)wbk.Worksheets["1"];
				
				Console.WriteLine(ws.Name);
				
				for (int i = 0; i < ws.Rows.Count; i++) {
					var cell = (Excel.Range)ws.Cells[i+1, 1];
					var value = cell.Value;
					
					if (value == null)
						break;
					
					string text = value.ToString();
					
					Console.WriteLine(text);
					
					list.Add(text);
				}
			}
			
			if (excel != null)
				excel.Quit();
			
			return list;
		}
		
		public static void ExportToWord(List<string> list)
        {
            Word.ApplicationClass word = null;

      		Word.Document doc = null;

            object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */
            
            try
            {
                word = new Word.ApplicationClass();
                word.Visible = true;
                doc = word.Documents.Add();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            
            if (word != null && doc != null)
            {
                Word.Table newTable;
                Word.Range wrdRng = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                
                newTable = doc.Tables.Add(wrdRng, list.Count, 1);
                newTable.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                newTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                newTable.AllowAutoFit = true;

                int i = 0;
                
                foreach (string s in list)
                {
                    newTable.Cell(i+1, 1).Range.Text = s;
                    i++;
                }

                doc.SaveAs(@"C:\1.docx");
            }
            
            if (word != null)
            	word.Quit();
        }
	}
}
