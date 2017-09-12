using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AgroInvestParsersLib
{
    class DetailedParser
    {

        int AreaId = 2;

        public void Parse(string path)
        {
            Application ObjExcel = new Application();
            Workbook ObjWorkBook = null;
            try
            {
                ObjWorkBook = ObjExcel.Workbooks.Open(path, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                var sourceSheet = (Worksheet)ObjWorkBook.Sheets["O-M"];
                var startDateRange = "LJ46";
                var endDateRange = "SQ46";
                var DateRange = sourceSheet.get_Range(startDateRange, endDateRange);

                string month=null;
                foreach (Range DateCell in DateRange.Cells)
                {
                    if (!(DateCell.Value is double))
                    {
                        month = "." + DateCell.Value.Substring(0, DateCell.Value.Length - 1);
                        month = month.Insert(4,"20");
                    }
                    else
                    {
                        AddCucArea(ObjWorkBook, DateCell.Column, month);
                        AddTomatoArea(ObjWorkBook, DateCell.Column, month);
                    }
                }
                ObjWorkBook.Close(true);
            }
            catch (Exception e){
                Console.WriteLine(e);
                Console.Read();
            }
            finally
            {
                ObjExcel.Quit();
            }
        }

        void AddCucArea(Workbook book, int j, string month)
        {
            var sourceSheet = (Worksheet)book.Sheets["O-M"];
            var targetSheet = (Worksheet)book.Sheets["Area"];

            targetSheet.Cells[1, 1].Value = "AreaId";
            targetSheet.Cells[1, 2].Value = "Культура";
            targetSheet.Cells[1, 3].Value = "Сорт";
            targetSheet.Cells[1, 4].Value = "Дата";
            targetSheet.Cells[1, 5].Value = "Сбор";

            var start = 98;
            var end = 102;
            for (var i = start; i <= end; i++)
            {
                var cult = sourceSheet.Cells[i, 1].Value;
                var date = sourceSheet.Cells[46, j].Value;
                var value = sourceSheet.Cells[i, j].Value;

                targetSheet.Cells[AreaId, 1].Value = AreaId;
                targetSheet.Cells[AreaId, 2].Value = "ОГУРЦЫ / CUCUMBER";
                targetSheet.Cells[AreaId, 3].Value = cult;
                if (date < 10)
                    targetSheet.Cells[AreaId, 4].Value = "0" + date + month;
                else
                    targetSheet.Cells[AreaId, 4].Value = date + month;

                targetSheet.Cells[AreaId, 5].Value = value;
                AreaId++;
            }
        }

        void AddTomatoArea(Workbook book, int j, string month)
        {
            var sourceSheet = (Worksheet)book.Sheets["Т-М"];
            var targetSheet = (Worksheet)book.Sheets["Area"];

            targetSheet.Cells[1, 1].Value = "AreaId";
            targetSheet.Cells[1, 2].Value = "Культура";
            targetSheet.Cells[1, 3].Value = "Сорт";
            targetSheet.Cells[1, 4].Value = "Дата";
            targetSheet.Cells[1, 5].Value = "Сбор";

            var start = 101;
            var end = 104;
            for (var i = start; i <= end; i++)
            {
                var cult = sourceSheet.Cells[i, 1].Value;
                var date = sourceSheet.Cells[51, j].Value;
                var value = sourceSheet.Cells[i, j].Value;

                targetSheet.Cells[AreaId, 1].Value = AreaId;
                targetSheet.Cells[AreaId, 2].Value = "ТОМАТЫ / TOMATO";
                targetSheet.Cells[AreaId, 3].Value = cult;
                if (date < 10)
                    targetSheet.Cells[AreaId, 4].Value = "0" + date + month;
                else
                    targetSheet.Cells[AreaId, 4].Value = date + month;

                targetSheet.Cells[AreaId, 5].Value = value;
                AreaId++;
            }
        }
    }
}
