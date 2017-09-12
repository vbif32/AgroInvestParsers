using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AgroInvestParsersLib
{
    class GatherParser
    {
            int GatherId = 1;

            public void Parse(string path)
            {
                Application ObjExcel = new Application();
                Workbook ObjWorkBook = null;
                try
                {
                    ObjWorkBook = ObjExcel.Workbooks.Open(path, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                    var sourceSheet = (Worksheet)ObjWorkBook.Sheets["Сбор"];
                    var targetSheet = (Worksheet)ObjWorkBook.Sheets["Gather"];

                    var entry = new string[8]
                    {
                        "GatherId",
                        "Date",
                        "Amount",
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
                        
                        row = FillProductHierachy(sourceSheet, row, ref entry);
                        if (row == 0)
                            break;
                        FillDateAndAmount(sourceSheet, targetSheet, row, ref entry);
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
                    var value = hierarcyCell.Value;

                    if (value == null)
                        return 0;

                    switch (level)
                    {
                        case 0:
                            entry[3] = value;
                            entry[4] = "";
                            entry[5] = "";
                            entry[6] = "";
                            entry[7] = "";
                            row++;
                            break;
                        case 2:
                            entry[4] = value;
                            entry[5] = "";
                            entry[6] = "";
                            entry[7] = "";
                            row++;
                            break;
                        case 4:
                            entry[5] = value;
                            entry[6] = "";
                            entry[7] = "";
                            row++;
                            break;
                        case 6:
                            entry[6] = value;
                            entry[7] = "";
                            row++;
                            break;
                        case 8:
                            entry[7] = value;
                            return row;
                        default:
                            break;
                    }
                }
            }

            public void FillDateAndAmount(Worksheet sourceSheet, Worksheet targetSheet, int row,ref string[] entry)
            {
                var startDateRange = "B4";
                var endDateRange = "EW4";
                var DateRange = sourceSheet.get_Range(startDateRange, endDateRange);

                var day = "";
                var month = "";
                var year = 2017;
                foreach (Range DateCell in DateRange.Cells)
                {
                    var dcValue = DateCell.Value;
                    if (dcValue is string)
                    {
                        switch (DateCell.Value)
                        {
                            case "Март/Mar":
                                month = "03";
                                break;
                            case "Апрель/Apr":
                                month = "04";
                                break;
                            case "Май/May":
                                month = "05";
                                break;
                            case "Июль/Jun":
                                month = "06";
                                break;
                            case "Июль/Jul":
                                month = "07";
                                break;
                            case "Август/Aug":
                                month = "08";
                                break;
                            default:
                                break;
                        }
                        continue;
                    }
                    if (dcValue is double)
                    {
                        if (dcValue < 10)
                            day = "0" + dcValue;
                        else
                            day = dcValue.ToString();


                        entry[1] = $"{day}.{month}.{year}";
                        entry[2] = sourceSheet.Cells[row, DateCell.Column].Value?.ToString();
                        WriteEntry(targetSheet, ref entry);
                    }
                }
            }

            void WriteEntry(Worksheet targetSheet, ref string[] entry)
            {
                targetSheet.Cells[GatherId, 1].Value = entry[0];
                targetSheet.Cells[GatherId, 2].Value = entry[1];
                targetSheet.Cells[GatherId, 3].Value = entry[2];
                targetSheet.Cells[GatherId, 4].Value = entry[3];
                targetSheet.Cells[GatherId, 5].Value = entry[4];
                targetSheet.Cells[GatherId, 6].Value = entry[5];
                targetSheet.Cells[GatherId, 7].Value = entry[6];
                targetSheet.Cells[GatherId, 8].Value = entry[7];
                Console.WriteLine(entry[0] + " " + entry[1] + " " + entry[2] + " " + entry[3] + " " + entry[4] + " " + entry[5] + " " + entry[6] + " " + entry[7]);
                GatherId++;
                entry[0]= GatherId.ToString();
            }
        }
    }
