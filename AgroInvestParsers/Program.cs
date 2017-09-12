using System;
using AgroInvestParsersLib;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace AgroInvestParsers
{
    class Program
    {
        static void Main(string[] args)
        {
            var logsPath = @"C:\TEMP\logs.txt";
            if (!File.Exists(logsPath))
                File.Create(logsPath);
            //var sprayingPath =
            //    @"C:\TEMP\Мониторинг теплиц\Мониторинг борьбы с вредителями\Опрыскивания_вредители_болезни\Опрыскивания\";
            var climatePath =
                @"C:\TEMP\Мониторинг теплиц\Мониторинг климата\";
            var climateVipPath =
                @"C:\TEMP\Мониторинг теплиц\Мониторинг климата\Уточненные VIP\";
            var feedingPath =
                @"C:\TEMP\Мониторинг теплиц\Мониторинг питания растений\";
            var feedingVipPath =
                @"C:\TEMP\Мониторинг теплиц\Мониторинг питания растений\Уточненные VIP\";
            var damagePath =
                @"C:\TEMP\Мониторинг теплиц\Мониторинг повреждений\";
            var CareTecnologyPath =
                @"C:\TEMP\Мониторинг теплиц\Мониторинг технологий ухода\";
            var indicatorsPath =
                @"C:\TEMP\Мониторинг теплиц\Показатели\";

            //// Spraying
            //try
            //{
            //    var start = $"{DateTime.Now:G} START spraying";
            //    Console.WriteLine(start);
            //    File.AppendAllText(logsPath,start);

            //    new SprayingParser().ParseExcel(sprayingPath);

            //    var end = $"{DateTime.Now:G} END spraying";
            //    Console.WriteLine(end);
            //    File.AppendAllText(logsPath, end);
            //}
            //catch (Exception e)
            //{
            //    var err = $"{DateTime.Now:G} {e}";
            //    Console.WriteLine(err);
            //    File.AppendAllText(logsPath, err);
            //}

            // Climate
            try
            {
                var start = $"{DateTime.Now:G} START climate\n";
                Console.WriteLine(start);
                File.AppendAllText(logsPath, start);

                foreach (var directory in Directory.GetDirectories(climatePath))
                    new MonitoringParser().ParseTxt(directory);

                var end = $"{DateTime.Now:G} END climate\n";
                Console.WriteLine(end);
                File.AppendAllText(logsPath, end);
            }
            catch (Exception e)
            {
                var err = $"{DateTime.Now:G} {e}";
                File.AppendAllText(logsPath, err);
                Console.WriteLine(err);
                Console.Beep();
                Console.Read();
            }

            // Climate VIP
            try
            {
                var start = $"{DateTime.Now:G} START climate VIP\n";
                Console.WriteLine(start);
                File.AppendAllText(logsPath, start);

                foreach (var directory in Directory.GetDirectories(climateVipPath))
                    new MonitoringParser().ParseTxt(directory);

                var end = $"{DateTime.Now:G} END climate VIP\n";
                Console.WriteLine(end);
                File.AppendAllText(logsPath, end);
            }
            catch (Exception e)
            {
                var err = $"{DateTime.Now:G} {e}";
                File.AppendAllText(logsPath, err);
                Console.WriteLine(err);
                Console.Beep();
                Console.Read();
            }

            // Feeding
            try
            {
                var start = $"{DateTime.Now:G} START feeding\n";
                Console.WriteLine(start);
                File.AppendAllText(logsPath, start);

                foreach (var directory in Directory.GetDirectories(feedingPath))
                    new MonitoringParser().ParseTxt(directory);
                var end = $"{DateTime.Now:G} END feeding\n";
                Console.WriteLine(end);
                File.AppendAllText(logsPath, end);
            }
            catch (Exception e)
            {
                var err = $"{DateTime.Now:G} {e}";
                File.AppendAllText(logsPath, err);
                Console.WriteLine(err);
                Console.Beep();
                Console.Read();
            }

            // Feeding VIP
            try
            {
                var start = $"{DateTime.Now:G} START feeding VIP\n";
                Console.WriteLine(start);
                File.AppendAllText(logsPath, start);

                foreach (var directory in Directory.GetDirectories(feedingVipPath))
                    new MonitoringParser().ParseTxt(directory);
                var end = $"{DateTime.Now:G} END feeding VIP\n";
                Console.WriteLine(end);
                File.AppendAllText(logsPath, end);
            }
            catch (Exception e)
            {
                var err = $"{DateTime.Now:G} {e}";
                File.AppendAllText(logsPath, err);
                Console.WriteLine(err);
                Console.Beep();
                Console.Read();
            }

            Console.WriteLine("End");
            Console.Read();
        }
    }
}
