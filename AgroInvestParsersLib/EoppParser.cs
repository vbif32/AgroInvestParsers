using Microsoft.Office.Interop.Excel;

namespace AgroInvestParsersLib
{
    class EoppParser
    {
        static long HarvestId = 2;
        static long PriceId = 2;
        static long RevenueId = 2;
        static long AreaId = 2;
        static long ProductivityId = 2;

        public void Parse(string path)
        {
            Application ObjExcel = new Application();
            Workbook ObjWorkBook = ObjExcel.Workbooks.Open(path, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            var sourceSheet = (Worksheet)ObjWorkBook.Sheets[1];

            var startDateRange = "B2";
            var endDateRange = "SQ2";
            var DateRange = sourceSheet.get_Range(startDateRange, endDateRange);

            foreach (Range DateCell in DateRange.Cells)
            {
                if (DateCell.Value is string)
                {
                    AddHarvest(ObjWorkBook, DateCell.Column);
                    AddPrice(ObjWorkBook, DateCell.Column);
                    AddRevenue(ObjWorkBook, DateCell.Column);
                    AddArea(ObjWorkBook, DateCell.Column);
                    AddProductivity(ObjWorkBook, DateCell.Column);
                }
            }
            ObjWorkBook.Close(true);
            ObjExcel.Quit();
        }

        void AddHarvest(Workbook book, int j)
        {
            var sourceSheet = (Worksheet)book.Sheets[1];
            var targetSheet = (Worksheet)book.Sheets[2];

            targetSheet.Cells[1, 1].Value = "HarvestId";
            targetSheet.Cells[1, 2].Value = "Культура";
            targetSheet.Cells[1, 3].Value = "Дата";
            targetSheet.Cells[1, 4].Value = "Сбор";

            var start = 5;
            var end = 11;
            for (var i = start; i <= end; i++)
            {
                var cult = sourceSheet.Cells[i, 1].Value;
                var date = sourceSheet.Cells[2, j].Value;
                date = date.Substring(0, date.Length - 1);
                var value = sourceSheet.Cells[i, j].Value;

                targetSheet.Cells[HarvestId, 1].Value = HarvestId;
                targetSheet.Cells[HarvestId, 2].Value = cult;
                targetSheet.Cells[HarvestId, 3].Value = date;
                targetSheet.Cells[HarvestId, 4].Value = value;
                HarvestId++;
            }
        }

        void AddPrice(Workbook book, int j)
        {
            var sourceSheet = (Worksheet)book.Sheets[1];
            var targetSheet = (Worksheet)book.Sheets[3];

            targetSheet.Cells[1, 1].Value = "PriceId";
            targetSheet.Cells[1, 2].Value = "Культура";
            targetSheet.Cells[1, 3].Value = "Дата";
            targetSheet.Cells[1, 4].Value = "Средняя цена";

            var start = 14;
            var end = 17;
            for (var i = start; i <= end; i++)
            {
                var cult = sourceSheet.Cells[i, 1].Value;
                var date = sourceSheet.Cells[2, j].Value;
                date = date.Substring(0, date.Length - 1);
                var value = sourceSheet.Cells[i, j].Value;

                targetSheet.Cells[PriceId, 1].Value = PriceId;
                targetSheet.Cells[PriceId, 2].Value = cult;
                targetSheet.Cells[PriceId, 3].Value = date;
                targetSheet.Cells[PriceId, 4].Value = value;
                PriceId++;
            }

        }

        void AddRevenue(Workbook book, int j)
        {
            var sourceSheet = (Worksheet)book.Sheets[1];
            var targetSheet = (Worksheet)book.Sheets[4];

            targetSheet.Cells[1, 1].Value = "RevenueId";
            targetSheet.Cells[1, 2].Value = "Культура";
            targetSheet.Cells[1, 3].Value = "Дата";
            targetSheet.Cells[1, 4].Value = "Доход";

            var start = 20;
            var end = 22;
            for (var i = start; i <= end; i++)
            {
                var cult = sourceSheet.Cells[i, 1].Value;
                var date = sourceSheet.Cells[2, j].Value;
                date = date.Substring(0, date.Length - 1);
                var value = sourceSheet.Cells[i, j].Value;

                targetSheet.Cells[RevenueId, 1].Value = RevenueId;
                targetSheet.Cells[RevenueId, 2].Value = cult;
                targetSheet.Cells[RevenueId, 3].Value = date;
                targetSheet.Cells[RevenueId, 4].Value = value;
                RevenueId++;
            }
        }

        void AddArea(Workbook book, int j)
        {
            var sourceSheet = (Worksheet)book.Sheets[1];
            var targetSheet = (Worksheet)book.Sheets[5];

            targetSheet.Cells[1, 1].Value = "AreaId";
            targetSheet.Cells[1, 2].Value = "Культура";
            targetSheet.Cells[1, 3].Value = "Дата";
            targetSheet.Cells[1, 4].Value = "Площадь посева";

            var start = 25;
            var end = 27;
            for (var i = start; i <= end; i++)
            {
                var cult = sourceSheet.Cells[i, 1].Value;
                var date = sourceSheet.Cells[2, j].Value;
                date = date.Substring(0, date.Length - 1);
                var value = sourceSheet.Cells[i, j].Value;

                targetSheet.Cells[AreaId, 1].Value = AreaId;
                targetSheet.Cells[AreaId, 2].Value = cult;
                targetSheet.Cells[AreaId, 3].Value = date;
                targetSheet.Cells[AreaId, 4].Value = value;
                AreaId++;
            }
        }

        void AddProductivity(Workbook book, int j)
        {
            var harvestSheet = (Worksheet)book.Sheets[2];
            var areaSheet = (Worksheet)book.Sheets[5];
            var targetSheet = (Worksheet)book.Sheets[6];

            targetSheet.Cells[1, 1].Value = "ProductivityId";
            targetSheet.Cells[1, 2].Value = "Культура";
            targetSheet.Cells[1, 3].Value = "Дата";
            targetSheet.Cells[1, 4].Value = "Урожайность";

            var start = 0;
            var end = 2;
            for (var i = start; i <= end; i++)
            {
                var cult = harvestSheet.Cells[HarvestId - 7 + i, 2].Value; ;
                var harvest = harvestSheet.Cells[HarvestId - 7 + i, 4].Value;
                var area = areaSheet.Cells[AreaId - 3 + i, 4].Value;
                var date = harvestSheet.Cells[HarvestId - 7 + i, 3].Value;
                var value = harvest / area * 1000;

                targetSheet.Cells[ProductivityId, 1].Value = ProductivityId;
                targetSheet.Cells[ProductivityId, 2].Value = cult;
                targetSheet.Cells[ProductivityId, 3].Value = date;
                targetSheet.Cells[ProductivityId, 4].Value = value;
                ProductivityId++;
            }
        }
    }
}
