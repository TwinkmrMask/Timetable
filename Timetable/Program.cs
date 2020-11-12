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
    class Program
    {
        private static byte[] file;
        private static string fileFormat;
        private static Dictionary<int, string> couples = new Dictionary<int, string>();
      
        static string[] GetDate()
        {
            string[] date = DateTime.Today.ToString("d.MM.yyyy").Split('.');
            var ru = CultureInfo.GetCultureInfo("ru-RU");
            date[0] = (int.Parse(date[0]) + 1).ToString();
            date[1] = ru.DateTimeFormat.MonthGenitiveNames[int.Parse(date[1]) - 1];
            return date;
        }

        static void GetFile(out byte[] file, out string fileFormat)
        {
            file = default;
            fileFormat = default;
            string[] date = GetDate();
            List<string> formats = new List<string>()
            {
                ".xls",  
                ".xlsx", 
                ".pdf"
            };
            foreach (string format in formats) 
            {
                string domain = $"http://www.mgkit.ru/studentu/raspisanie-zanatij/РАСПИСАНИЕ%20{date[0]}%20{date[1]}%20{GetDate()[2]}{format}?attredirects=0&d=1";
                try
                {
                    WebClient wc = new WebClient();
                    file = wc.DownloadData(domain);
                    fileFormat = format;
                    break;
                }
                finally
                {
                    file = default;
                    fileFormat = default;
                }
            }
        }

        static void Timetable(in byte[] file, in string fileFormat)
        {
            if ((file != null) || (file != default))
            {
                dynamic wb = null;
                try
                {
                    if (fileFormat != ".pdf")
                    {
                        using (MemoryStream timetable = new MemoryStream(file))
                        {
                            if (fileFormat == ".xls")
                                wb = new HSSFWorkbook(timetable);

                            if (fileFormat == ".xlsx")
                            {
                                wb = new XSSFWorkbook(timetable);

                                ISheet sheet = wb.GetSheetAt(0);

                                for (int i = 16; i < 16 + 14; i++)
                                {
                                    var cell = sheet.GetRow(i);
                                    if (i % 2 != 0)
                                        couples.Add((i - 15) / 2, cell.GetCell(21).ToString().Replace("\n", " ").Trim());
                                }

                                if (couples.Values.Any(v => !string.IsNullOrWhiteSpace(v))) 
                                    foreach (KeyValuePair<int, string> _couples in couples)
                                    {
                                        Console.Write(_couples.Key + ".");
                                        Console.ForegroundColor = ConsoleColor.DarkMagenta;
                                        Console.WriteLine(_couples.Value);
                                }
                                else
                                {
                                    Console.ForegroundColor = ConsoleColor.Green;
                                    Console.WriteLine("\n      День\n Самостоятельной\n     Работы");
                                    Console.ResetColor();
                                }
                                
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
            GetFile(file: out file, fileFormat: out fileFormat);
            Timetable(file: in file, fileFormat: in fileFormat);
            Console.ReadKey(true);
        }
    }
}