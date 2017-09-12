using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AgroInvestParsersLib
{
    class CropParser
    {
        int id = 0;
        public void Parse()
        {
            var excel = new Application();
            var list = new List<string>();
            try
            {
                var paths = new[]
                {
                    @"C:\TEMP\Crop registration CUC GH 1.xls",
                    @"C:\TEMP\Crop registration TOM GH 1.xls",
                    @"C:\TEMP\Crop registration TOM GH 2 — 2017-2018.xls",
                    @"C:\TEMP\Crop registration TOM GH 2.xls",
                };
                var sheets = new[]
                {
                    "2016-2017",
                    "2016-2017",
                    "2017-2018",
                    "2016-2017",
                };
                var productTypes2 = new[]
                {
                    "ОГУРЦЫ / CUCUMBER",
                    "ТОМАТЫ / TOMATO",
                    "ТОМАТЫ / TOMATO",
                    "ТОМАТЫ / TOMATO",
                };
                var ghNumbers = new[]
                {
                    1,
                    1,
                    2,
                    2,
                };
                var startDates = new[]
                {
                    "1.08.2016",
                    "13.04.2016",
                    "1.08.2016",
                    "1.08.2016",
                };
                for (var i = 0; i < paths.Length; i++)
                {
                    var book = excel.Workbooks.Open(paths[i], 0, true, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    var sheet = book.Sheets[sheets[i]];
                    var prodType2 = productTypes2[i];
                    var ghNumber = ghNumbers[i];
                    var startDate = DateTime.Parse(startDates[i]);
                    var path = @"C:\TEMP\" + book.Name + ".csv";
                    Console.WriteLine(book.Name);
                    for (var row = 11;; row++)
                    {
                        Console.Write(row);
                        var date = startDate;
                        Range paramCell = sheet.Cells[row, 1];
                        Range prodType4Cell = sheet.Cells[row, 38];
                        string param;
                        string prodType4;
                        if (paramCell.MergeCells)
                        {
                            var tmp = paramCell.MergeArea.Value;
                            param = tmp[1,1];
                        }
                        else
                            param = paramCell.Value;
                        if (prodType4Cell.MergeCells)
                        {
                            var tmp = prodType4Cell.MergeArea.Value;
                            prodType4 = tmp[1, 1];
                        }
                        else
                            prodType4 = prodType4Cell.Value;

                        if (prodType4 is null || String.IsNullOrEmpty(prodType4) || param is null || String.IsNullOrEmpty(param))
                            break;

                        for (var column = 39;; column++)
                        {
                            Console.Write(" " + column);
                            string weekNumber = sheet.Cells[4, column].Value?.ToString();
                            if (weekNumber is null || String.IsNullOrEmpty(weekNumber))
                                break;
                            double.TryParse(sheet.Cells[row, column].Value?.ToString(), out double value);
                            list.Add($"{id};{param};{ghNumber};{prodType2};{prodType4};{date:dd.MM.yyyy};{Math.Round(value, 3)}");
                            id++;
                            date = date.AddDays(7);
                        }
                        Console.WriteLine(" ok");
                    }
                    book.Close(true);
                    WriteToCSV(path, list);
                    list.Clear();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            finally
            {
                excel.Quit();
                Console.WriteLine("End");
            }
        }

        void WriteToCSV(string path,List<string> list)
        {
            var text = list.Aggregate((current, next) => current + "\r\n" + next);
            System.IO.File.WriteAllText(path, text);
        }

    }
}
