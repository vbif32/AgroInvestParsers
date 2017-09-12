using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace AgroInvestParsersLib
{
    public class MonitoringParser :ParserBase
    {
        protected override void ParseBook(Workbook book)
        {
            throw new NotImplementedException();
        }

        protected override void ParseFile(string path)
        {
            var list = new List<string> {"Id;ParamName;GhNumber;Branch;Gr;DateTime;Value"};
            using (var sr = new StreamReader(new FileStream(path, FileMode.Open, FileAccess.Read)))
            {
                var param = path.Substring(path.LastIndexOf('\\') + 1,
                    path.LastIndexOf('.') - 1 - path.LastIndexOf('\\'));

                const string pattern = @"\d";
                var r = new Regex(pattern);
                var m = r.Match(path);
                var gh = m.Value;
                var gr = "";
                var branch = "";
                string s;
                while ((s = sr.ReadLine()) != null)
                {
                    if (string.IsNullOrEmpty(s))
                        continue;
                    if (s.Contains("Gr "))
                    {
                        gr = s.Substring(s.IndexOf("Gr ")).Trim();
                        if (!int.TryParse(gr.Substring(gr.Length - 1), out int tmp)) continue;
                        if (tmp < 3)
                            branch = "1";
                        else if (tmp < 5)
                            branch = "2";
                        else if (tmp < 7)
                            branch = "3";
                        else if (tmp < 9)
                            branch = "4";
                    }
                    else if (char.IsDigit(s, 0))
                    {
                        var ss = s.Split();
                        var date = ss[0];
                        var time = ss[1];
                        var dateTime = date + " " + time;
                        var value = ss[2];
                        var entry = $"{Id};{param};{gh};{branch};{gr};{dateTime};{date};{time};{value}";
                        list.Add(entry);
                        Console.WriteLine(entry);
                        Id++;
                    }
                }
                Console.WriteLine("Start writing csv");
                WriteToCsv(Path.Combine(CsvDirectory, param + ".csv"), list);
                Console.WriteLine("End writing csv");
                list.Clear();
            }
        }
    }
}