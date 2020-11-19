using System;
using System.Collections.Generic;
using System.Net;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Globalization;
using System.Linq;

namespace Timetable
{
    class Program // The processing logic schedules
    {
        private static byte[] file;
        private static string fileFormat;
        private static string[] day;
        private static Dictionary<int, string> couples = new Dictionary<int, string>();
        static string[] GetDate(dynamic date)
        {
            var ru = CultureInfo.GetCultureInfo("ru-RU");
            day = DateTime.Today.ToString("d.MM.yyyy").Split('.');
            day[1] = ru.DateTimeFormat.MonthGenitiveNames[int.Parse(day[1]) - 1];
            Console.WriteLine($"{day[0]} {day[1]} {day[2]}");
            if (date == "today")
                return day;
            if (date == "tomorrow")
                day[0] = (int.Parse(day[0]) + 1).ToString();
            else
                day = date.Split('.');

            return day;

        }
        static void GetFile(out byte[] file, out string fileFormat, string day)
        {
            file = default;
            fileFormat = default;
            string[] date = GetDate(day);
            List<string> formats = new List<string>() { ".xls", ".xlsx", ".pdf" };
            foreach (string format in formats)
            {
                string domain = $"http://www.mgkit.ru/studentu/raspisanie-zanatij/РАСПИСАНИЕ%20{date[0]}%20{date[1]}%20{date[2]}{format}?attredirects=0&d=1";
                try
                {
                    WebClient wc = new WebClient();
                    file = wc.DownloadData(domain);
                    fileFormat = format;
                    break;
                }
                catch
                {
                    file = default;
                    fileFormat = default;
                    //<-------------------------------->
                    Console.WriteLine(domain);
                    Console.ForegroundColor = ConsoleColor.Cyan;
                    Console.WriteLine("http://www.mgkit.ru/studentu/raspisanie-zanatij/РАСПИСАНИЕ%20%2020%20ноября%202020.xlsx?attredirects=0&d=1" == domain);
                    Console.ResetColor();
                }
            }
            Console.WriteLine("http://www.mgkit.ru/studentu/raspisanie-zanatij/РАСПИСАНИЕ%20%2020%20ноября%202020.xlsx?attredirects=0&d=1");
            if (file == default)
            {
                Console.WriteLine("There is no schedule for tomorrow yet, but there is one for today");
                GetFile(out file, out fileFormat, "today");
            }
        }
        static void Timetable(in byte[] file, in string fileFormat)
        {
            if ((file != null) || (file != default))
            {
                IWorkbook wb = null;
                try
                {
                    if (fileFormat != ".pdf")
                    {
                        using (MemoryStream timetable = new MemoryStream(file))
                        {
                            if (fileFormat == ".xls")
                                wb = new HSSFWorkbook(timetable);

                            if (fileFormat == ".xlsx")
                                wb = new XSSFWorkbook(timetable);

                            ISheet sheet = wb.GetSheetAt(0);

                            for (int i = 16; i < 16 + 14; i++)
                            {
                                var cell = sheet.GetRow(i);

                                if (cell.GetCell(21) == null)
                                {
                                    if (i % 2 != 0)
                                        couples.Add((i - 15) / 2, null);                                    
                                    i++;
                                }

                                if (i % 2 != 0)
                                    couples.Add((i - 15) / 2, cell.GetCell(21).StringCellValue.Replace("\n", " ").Trim());

                            }
                                if (couples.Values.Any(v => !string.IsNullOrWhiteSpace(v)))
                                {
                                 //using (DataBase data = new DataBase())
                                 //{
                                 //string[] date = GetDate("tomorrow");
                                foreach (KeyValuePair<int, string> _couples in couples)
                                    {
                                        Console.Write(_couples.Key + ". ");
                                        Console.ForegroundColor = ConsoleColor.Cyan;
                                        Console.WriteLine(_couples.Value);
                                        Console.ResetColor();
                                    }
                                    //}
                                }
                                else
                                {
                                    Console.ForegroundColor = ConsoleColor.Green;
                                    Console.WriteLine("\n      День\n Самостоятельной\n     Работы");
                                    Console.ResetColor();
                                }
                            
                        }
                    }
                    else
                    {
                        Console.WriteLine("Oops, i don't know how to handle this format yet");
                    }
                }
                finally
                {
                    wb.Close();
                }
            }
            else
            {
                Console.WriteLine("File is empty");
            }
        }
        static void Main(string[] args)
        {
            GetFile(file: out file, fileFormat: out fileFormat, day: "tomorrow");
            Timetable(file: in file, fileFormat: in fileFormat);
            Console.ReadKey(true);
        }
    }
}
