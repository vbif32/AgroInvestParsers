using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AgroInvestParsersLib
{
    class TeacherParser
    {
        int row = 1;

        public void Parse(string path)
        {
            Application ObjExcel = new Application();
            Workbook ObjWorkBook = null;
            try
            {
                ObjWorkBook = ObjExcel.Workbooks.Open(path, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                Worksheet targetSheet;
                try
                {
                    targetSheet = ObjWorkBook.Sheets["Teachers"];
                }
                catch
                {
                    targetSheet = ObjWorkBook.Sheets.Add();
                    targetSheet.Name = "Teachers";
                }
                
                int[] pIndex = new int[15] { 1, 4, 6, 7, 8, 9, 11, 13, 15, 17, 19, 22, 24, 26, 33 };
                var entry = new string[34]
                {
                    "teacher",
                    "p1",
                    "napravl",
                    "sub_name",
                    "p2",
                    "sem",
                    "p3",
                    "p4",
                    "p5",
                    "p6",
                    "lec",
                    "p7",
                    "labs",
                    "p8",
                    "pr",
                    "p9",
                    "exam",
                    "p10",
                    "zach",
                    "p11",
                    "kp",
                    "kons",
                    "p12",
                    "vkr",
                    "p13",
                    "gek",
                    "p14",
                    "gak",
                    "practice",
                    "ruk",
                    "asp",
                    "konsult",
                    "sum",
                    "p15"
                };
                WriteEntry(targetSheet, entry);

                var length = ObjWorkBook.Sheets.Count - 3;
                //var length = 4;
                for (int i = 5; i <= length; i++)
                {
                    var sourceSheet = (Worksheet)ObjWorkBook.Sheets[i];
                    ParseSheet(sourceSheet,targetSheet, ref entry);
                }

                ObjWorkBook.Close(true);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            finally
            {
                Console.Read();
                ObjExcel.Quit();
            }
        }

        void WriteEntry(Worksheet worksheet, string[] entry)
        {
            var log = "";
            for (int i = 0; i < entry.Length; i++)
            {
                worksheet.Cells[row, i + 1] = entry[i];
                log += entry[i] + ";";
            }
            Console.WriteLine(log);
            row++;
        }

        void ParseSheet(Worksheet worksheet, Worksheet targetSheet, ref string[] entry)
        {
            entry[0] = worksheet.Name;

            int headerRow = FindHeaderRow(worksheet);
            for (int row = headerRow + 1; ; row++)
            {
                if (worksheet.Cells[row, 1].Value == null || worksheet.Cells[row, 1].Value.Contains("Всего"))
                    return;

                int[] pIndex = new int[15] { 1, 4, 6, 7, 8, 9, 11, 13, 15, 17, 19, 22, 24, 26, 33 };

                int j = 1;
                for (int column = 1; ; column++)
                {
                    var header = worksheet.Cells[headerRow, column].Value;
                    var headeAdd = worksheet.Cells[headerRow, column].Address;
                    if (worksheet.Cells[headerRow, column].Value == null)
                        continue;

                    var value = worksheet.Cells[row, column].Value;

                    if (!pIndex.Contains(j) && value == null)
                    {
                        worksheet.Cells[row, column].Value = 0;
                        value = 0;
                    }
                        

                    entry[j] = value?.ToString();
                    j++;
                    if (j == entry.Length)
                        break;
                }

                WriteEntry(targetSheet, entry);
            }
        }
        int FindHeaderRow(Worksheet worksheet)
        {
            for (int i = 1; ; i++)
            {
                var value = worksheet.Cells[i, 1].Value;
                if (value == "p1")
                    return i;
            }
        }
    }
}

