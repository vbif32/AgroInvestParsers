using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AgroInvestParsersLib
{
    class PriceParser
    {
        int PriceId = 1;

        public void Parse(string path)
        {
            Application ObjExcel = new Application();
            Workbook ObjWorkBook = null;
            try
            {
                ObjWorkBook = ObjExcel.Workbooks.Open(path, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                var sourceSheetKg = (Worksheet)ObjWorkBook.Sheets["Продажи кг"];
                var sourceSheetRub = (Worksheet)ObjWorkBook.Sheets["Продажи руб"];
                var targetSheet = (Worksheet)ObjWorkBook.Sheets["Price"];

                var entry = new string[9]
                {
                    "PriceId",
                    "Date",
                    "SalesKg",
                    "SalesRub",
                    "Product_type_1",
                    "Product_type_2",
                    "Product_type_3",
                    "Product_type_4",
                    "Product_type_5",
                };
                WriteEntry(targetSheet, ref entry);

                var row = 6;
                while (true)
                {
                        
                    row = FillProductHierachy(sourceSheetKg, row, ref entry);
                    if (row == 0)
                        break;
                    FillDateAndAmount(sourceSheetKg, sourceSheetRub, targetSheet, row, ref entry);
                    row++;
                }
                ObjWorkBook.Close(true);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Console.Read();
            }
            finally
            {
                ObjExcel.Quit();
            }
        }

        public int FillProductHierachy(Worksheet sourceSheet, int row, ref string[] entry)
        {
            while (true)
            {
                var hierarcyCell = sourceSheet.Cells[row, 1] as Range;
                var level = hierarcyCell.IndentLevel;
                var nextLevel = sourceSheet.Cells[row+1, 1].IndentLevel;
                var value = hierarcyCell.Value;

                if (value == null || row == 112)
                    return 0;

                if ((level >= nextLevel) && (level !=8) && (level !=0) )
                {
                    for(var i = 4; i < entry.Length; i++)
                    {
                        if (entry[i] == "")
                        {
                            entry[8] = value;
                            return row;
                        }
                    }
                }



                switch (level)
                {
                    case 0:
                        entry[4] = value;
                        entry[5] = "";
                        entry[6] = "";
                        entry[7] = "";
                        entry[8] = "";
                        row++;
                        break;
                    case 2:
                        entry[5] = value;
                        entry[6] = "";
                        entry[7] = "";
                        entry[8] = "";
                        row++;
                        break;
                    case 4:
                        entry[6] = value;
                        entry[7] = "";
                        entry[8] = "";
                        row++;
                        break;
                    case 6:
                        entry[7] = value;
                        entry[8] = "";
                        row++;
                        break;
                    case 8:
                        entry[8] = value;
                        return row;
                    default:
                        break;
                }
            }
        }

        public void FillDateAndAmount(Worksheet sourceSheetKg, Worksheet sourceSheetRub, Worksheet targetSheet, int row,ref string[] entry)
        {
            var startDateRange = "B4";
            var endDateRange = "IA4";
            var DateRange = sourceSheetKg.get_Range(startDateRange, endDateRange);

            foreach (Range DateCell in DateRange.Cells)
            {
                string dcValue = DateCell.Value;
                if (dcValue == null)
                    continue;
                if (dcValue.Contains(":"))
                {
                    var date = dcValue.Split(' ')[0];
                    entry[1] = date;
                    entry[2] = sourceSheetKg.Cells[row, DateCell.Column].Value?.ToString();
                    entry[3] = sourceSheetRub.Cells[row+1, DateCell.Column].Value?.ToString();
                    WriteEntry(targetSheet, ref entry);
                }
            }
        }

        void WriteEntry(Worksheet targetSheet, ref string[] entry)
        {
            targetSheet.Cells[PriceId, 1].Value = entry[0];
            targetSheet.Cells[PriceId, 2].Value = entry[1];
            targetSheet.Cells[PriceId, 3].Value = entry[2];
            targetSheet.Cells[PriceId, 4].Value = entry[3];
            targetSheet.Cells[PriceId, 5].Value = entry[4];
            targetSheet.Cells[PriceId, 6].Value = entry[5];
            targetSheet.Cells[PriceId, 7].Value = entry[6];
            targetSheet.Cells[PriceId, 8].Value = entry[7];
            targetSheet.Cells[PriceId, 9].Value = entry[8];
            Console.WriteLine(entry[0] + " " + entry[2] + " " + entry[3] + " " + entry[4] + " " + entry[5] + " " + entry[6] + " " + entry[7] + " " + entry[8]);
            PriceId++;
            entry[0]= PriceId.ToString();
        }
    }
}
