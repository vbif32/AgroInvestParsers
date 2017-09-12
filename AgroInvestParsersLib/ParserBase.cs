using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace AgroInvestParsersLib
{
    public abstract class ParserBase
    {
        protected long Id;
        protected string CsvDirectory;

        public void ParseExcel(string directory)
        {
            var paths = Directory.GetFiles(directory);
            CsvDirectory = GetOutputdirectory(directory);
            var excel = new Application();
            foreach (var path in paths)
            {
                var book = excel.Workbooks.Open(path, 0, true, 5, "", "", false, XlPlatform.xlWindows, "", true, false,
                    0, true, false, false);
                ParseBook(book);
            }
            excel.Quit();
        }
        public void ParseExcel(string[] paths)
        {
            var excel = new Application();
            foreach (var path in paths)
            {
                CsvDirectory = GetOutputdirectory(Path.GetDirectoryName(path));
                var book = excel.Workbooks.Open(path, 0, true, 5, "", "", false, XlPlatform.xlWindows, "", true, false,
                    0, true, false, false);
                ParseBook(book);
            }
            excel.Quit();
        }
        public void ParseTxt(string directory)
        {
            CsvDirectory = GetOutputdirectory(directory);
            var paths = Directory.GetFiles(directory, "*.txt");
            foreach (var path in paths)
                ParseFile(path);
        }
        public void ParseTxt(string[] paths)
        {
            foreach (var path in paths)
            {
                CsvDirectory = GetOutputdirectory(Path.GetDirectoryName(path));
                ParseFile(path);
            }
        }
        protected abstract void ParseBook(Workbook book);
        protected abstract void ParseFile(string path);


        public static string ToDate(int day, int month, int year)
        {
            var sday = day > 9 ? day.ToString() : "0" + day;
            var smonth = month > 9 ? month.ToString() : "0" + month;
            var syear = year > 1999 ? year.ToString() : "20" + year;
            return $"{sday}.{smonth}.{syear}";
        }
        public string GetOutputdirectory(string path)
        {
            return path.Insert(path.IndexOf("TEMP") + 5, @"CSV\");
        }
        public static void WriteToCsv(string path, IEnumerable<string> list)
        {
            try
            {
                var fileInfo = new FileInfo(path);
                if (!fileInfo.Exists)
                    Directory.CreateDirectory(fileInfo.Directory.FullName);
                File.WriteAllLines(path, list);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
