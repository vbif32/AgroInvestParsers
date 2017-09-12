using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Excel;

namespace AgroInvestParsersLib
{
    public class ButterflyParser : ParserBase
    {
        protected override void ParseBook(Workbook book)
        {
            ParseGather(book);
            ParseBudget(book);
        }

        protected override void ParseFile(string path)
        {
            throw new NotImplementedException();
        }

        private void ParseGather(_Workbook book)
        {
            var sheetNames = new[]
            {
                "Сбор2016",
                "Сбор2017"
            };
            foreach (var sheetName in sheetNames)
            {
                _Worksheet sheet = book.Sheets[sheetName];
                var list = new List<string>();
                var ptHier = new string[5];
                const string header = "Id;ProductType1;ProductType2;ProductType3;ProductType4;ProductType5;Date;Value";
                list.Add(header);

                const int startRow = 9;
                const int endRow = 99;
                const int startColumn = 2;
                const int endColumn = 357;

                for (var row = startRow; row <= endRow; row++)
                {
                    Console.WriteLine(row + " ");
                    Range pt = sheet.Cells[row, 1];
                    if (pt.IndentLevel == 8 || pt.Interior.ColorIndex == -4142)
                    {
                        ptHier[4] = pt.Value;
                    }
                    else if (pt.IndentLevel == 6)
                    {
                        ptHier[3] = pt.Value;
                        ptHier[4] = "";
                        continue;
                    }
                    else if (pt.IndentLevel == 4)
                    {
                        ptHier[2] = pt.Value;
                        ptHier[3] = "";
                        ptHier[4] = "";
                        continue;
                    }
                    else if (pt.IndentLevel == 2)
                    {
                        ptHier[1] = pt.Value;
                        ptHier[2] = "";
                        ptHier[3] = "";
                        ptHier[4] = "";
                        continue;
                    }
                    else if (pt.IndentLevel == 0)
                    {
                        ptHier[0] = pt.Value;
                        ptHier[1] = "";
                        ptHier[2] = "";
                        ptHier[3] = "";
                        ptHier[4] = "";
                        continue;
                    }
                    var month = 0;
                    var year = int.Parse(sheet.Name.Substring(sheetName.Length - 4));
                    for (var column = startColumn; column <= endColumn; column++)
                    {
                        Console.Write(column + " ");
                        Range dateCell = sheet.Cells[7, column];
                        var dateCellValue = dateCell.Value;
                        if (dateCellValue is null)
                            break;
                        if (dateCellValue is string)
                        {
                            month++;
                            continue;
                        }
                        var day = (int) dateCellValue;
                        var value = sheet.Cells[row, column].Value;
                        list.Add(
                            $"{Id};{ptHier[0]};{ptHier[1]};{ptHier[2]};{ptHier[3]};{ptHier[4]};{ToDate(day, month, year)};{value}");
                        Id++;
                    }
                    Console.WriteLine("ok");
                }
                WriteToCsv($@"{CsvDirectory}{book.Name}_{sheet.Name}.csv", list);
            }
        }
        private void ParseBudget(_Workbook book)
        {
            _Worksheet sheet = book.Sheets["СборБюдж"];
            var list = new List<string>();
            const string header = "Id;ProductType2;ProductType3;Date;Value";
            list.Add(header);

            const int startRow = 5;
            const int endRow = 11;
            const int startColumn = 3;
            const int endColumn = 13;

            for (var row = startRow; row <= endRow; row++)
            {
                var pt2 = row < 8 ? "ОГУРЦЫ / CUCUMBER" : "ТОМАТЫ / TOMATO";
                var pt3 = sheet.Cells[row, 2].Value;
                for (var column = startColumn; column <= endColumn; column++)
                {
                    var date = sheet.Cells[3, column].Value;
                    var value = (int) Math.Round(sheet.Cells[row, column].Value);
                    list.Add($"{Id};{pt2};{pt3};{date:dd.MM.yyyy};{value}");
                    Id++;
                }
            }
            WriteToCsv($@"C:\TEMP\{book.Name}_{sheet.Name}.csv", list);
        }
    }
}