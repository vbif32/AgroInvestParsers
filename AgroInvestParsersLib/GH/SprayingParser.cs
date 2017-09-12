using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace AgroInvestParsersLib
{
    public class SprayingParser : ParserBase
    {
        protected override void ParseBook(Workbook book)
        {
            ParseSpraying(book);
        }

        protected override void ParseFile(string path)
        {
            throw new NotImplementedException();
        }

        private void ParseSpraying(_Workbook book)
        {
            var sheetNames = new[]
            {
                "2016",
                "2017"
            };
            const string name = "Spraying";
            foreach (var sheetName in sheetNames)
            {
                _Worksheet sheet = book.Sheets[sheetName];
                var list = new List<string>();
                const string header = "Id;Date;GHNumber;Branch;Value";
                list.Add(header);

                const int startRow = 3;
                const int endRow = 370;

                for (var row = startRow; row <= endRow; row++)
                {
                    Console.Write(row);
                    var date = sheet.Cells[row, 2].Value?.ToString().Split()[0];
                    var ghNumber = sheet.Cells[row, 3].Value?.ToString();
                    if (date is null)
                        break;

                    for (var i = 1; i <= 4; i++)
                    {
                        var branch = i;
                        var value = sheet.Cells[row, i+3].Value;
                        list.Add($"\"{Id}\";\"{date}\";\"{ghNumber}\";\"{branch}\";\"{name}\";\"{value}\"");
                        Id++;
                    }
                    Id++;

                    Console.WriteLine(" ok");
                }
                WriteToCsv($@"C:\TEMP\{book.Name}_{sheet.Name}.csv", list);
            }
        }
    }
}
