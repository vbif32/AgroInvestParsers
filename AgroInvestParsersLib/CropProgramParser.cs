using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace AgroInvestParsersLib
{
    class CropProgramParser

    {
        private long _id;

        public void Parse()
        {
            var paths = new[]
            {
                @"C:\TEMP\Производственная программа 2017.xlsx"
            };
            var excel = new Application();
            foreach (var path in paths)
            {
                var book = excel.Workbooks.Open(path, 0, true, 5, "", "", false, XlPlatform.xlWindows, "", true, false,
                    0, true, false, false);
                ParseCropProgram(book);
            }
            excel.Quit();
        }


        private void ParseCropProgram(_Workbook book)
        {
            var sheetNames = new[]
            {
                "Т1-1",
                "Т1-2",
                "Т2"
            };
            foreach (var sheetName in sheetNames)
            {
                Console.WriteLine(sheetName);
                _Worksheet sheet = book.Sheets[sheetName];
                var list = new List<string>();
                const string header = "Id;GHNumber;Branch;ProductType2;ProductType3;Date;Name;Value";
                list.Add(header); 

                const int startRow = 8;
                const int endRow = 27;
                const int startColumn = 3;
                const int endColumn = 54;
                var date = DateTime.Parse("01.01.2017");
                var ghNumber = sheetName.ElementAt(1);
                string branch = null;
                var names = new [] { "Рассада", "Вегетация", "Плодоношение", "Ликвидация" };
                string value = null;
                for (var column = startColumn; column <= endColumn; column++)
                {
                    Console.WriteLine("     " + column + " ");
                    var pt2 = sheetName == sheetNames[0] ? "ОГУРЦЫ / CUCUMBER" : "ТОМАТЫ / TOMATO";
                    var values = new[]{"0", "0", "0", "0"};
                    for (var row = startRow; row <= endRow; row++)
                    {
                        switch (row)
                        {
                            case int r when r < 13:
                                branch = "4";
                                break;
                            case int r when r < 18:
                                branch = "3";
                                break;
                            case int r when r < 23:
                                branch = "2";
                                break;
                            case int r when r < 28:
                                branch = "1";
                                break;
                        }
                        string pt3 = sheet.Cells[row, 2].Value;
                        switch (pt3)
                        {
                            case string s when s.Contains("Отделение"):
                                pt3 = "";
                                break;
                        }
                        int color = sheet.Cells[row, column].Interior.ColorIndex;
                        switch (color)
                        {
                            case 36:
                                values[0] = "1";
                                break;
                            case 15:
                                values[1] = "1";
                                break;
                            case 14:
                                values[2] = "1";
                                break;
                            case 3:
                                values[3] = "1";
                                break;
                        }
                        for (var i = 0; i < 4; i++)
                        {
                            var entry = $"{_id};{ghNumber};{branch};{pt2};{pt3};{date:dd.MM.yyyy};{names[i]};{values[i]}";
                            list.Add(entry);
                            Console.WriteLine("         " + entry);
                            _id++;
                        }
                    }
                    date = date.AddDays(7);
                    Console.WriteLine("ok");
                }
                Console.WriteLine("Start writing csv");
                WriteToCsv($@"C:\TEMP\{book.Name}_{sheet.Name}.csv", list);
                Console.WriteLine("Stop writing csv");
            }
        }

        private static void WriteToCsv(string path, IEnumerable<string> list)
        {
            try
            {
                using (var sw = new StreamWriter(path,false, Encoding.Default))
                {
                    foreach (var line in list)
                        sw.WriteLine(line);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
