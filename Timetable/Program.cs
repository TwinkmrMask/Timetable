using System;
using System.Collections.Generic;
using System.Net;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace Timetable
{
    class Program
    {
        private static string path = default;
        private static Dictionary<int, string> couples = new Dictionary<int, string>();
        private static Dictionary<string, string> months = new Dictionary<string, string>()
        {
            ["01"] = "января",
            ["02"] = "февраля",
            ["03"] = "марта",
            ["04"] = "апреля",
            ["05"] = "мая",
            ["06"] = "июня",
            ["07"] = "июля",
            ["08"] = "августа",
            ["09"] = "сентября",
            ["10"] = "октября",
            ["11"] = "ноября",
            ["12"] = "декабря"
        };

        static string[] GetDate() => DateTime.Today.ToString("d.MM.yyyy").Split('.');

        static void GetFile(out string path)
        {
            path = default;
            Dictionary<string, bool> formats = new Dictionary<string, bool>()
            {
                [".xls"] = true,
                [".xlsx"] = true,
                [".pdf"] = true
            };

            try
            {
                WebClient wc = new WebClient();
                wc.DownloadFile(
                    $"http://www.mgkit.ru/studentu/raspisanie-zanatij/РАСПИСАНИЕ%20{GetDate()[0]}%20{months[GetDate()[1]]}%20{GetDate()[2]}.xls?attredirects=0&d=1",
                    $"C:\\Users\\user\\Downloads\\Расписание на {GetDate()[0]} {months[GetDate()[1]]}.xls"
                    );
            }
            catch (Exception)
            {
                formats[".xls"] = false;
            }
            try
            {
                WebClient wc = new WebClient();
                wc.DownloadFile(
                    $"http://www.mgkit.ru/studentu/raspisanie-zanatij/РАСПИСАНИЕ%20{GetDate()[0]}%20{months[GetDate()[1]]}%20{GetDate()[2]}.xlsx?attredirects=0&d=1",
                    $"C:\\Users\\user\\Downloads\\Расписание на {GetDate()[0]} {months[GetDate()[1]]}.xlsx"
                    );
            }
            catch (Exception)
            {
                formats[".xlsx"] = false;
            }
            try
            {
                WebClient wc = new WebClient();
                wc.DownloadFile(
                    $"http://www.mgkit.ru/studentu/raspisanie-zanatij/РАСПИСАНИЕ%20{GetDate()[0]}%20{months[GetDate()[1]]}%20{GetDate()[2]}.pdf?attredirects=0&d=1",
                    $"C:\\Users\\user\\Downloads\\Расписание на {GetDate()[0]} {months[GetDate()[1]]}.pdf"
                    );
            }
            catch (Exception)
            {
                formats[".pdf"] = false;
            }

            foreach (string format in formats.Keys)
            {
                if (formats[format] == true)
                    path = $"C:\\Users\\user\\Downloads\\Расписание на {GetDate()[0]} {months[GetDate()[1]]}{format}";
            }
        }

        static void GetTimetable(in string path)
        {
            if ((path != default) || (path != null))
            {
                
                if (path.Substring(path.Length - 4) == ".xls")
                {
                    HSSFWorkbook wb = default;
                    try
                    {
                        using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
                        {
                            wb = new HSSFWorkbook(file);
                        }
                        ISheet sheet = wb.GetSheetAt(0);
                        int i = 1;
                        var cell = sheet.GetRow(i);
                        Console.WriteLine(cell.GetCell(0).StringCellValue);
                        //while ((cell != null) && (cell.GetCell(0).StringCellValue != ""))
                        //{
                        //   
                        //}
                    }
                    finally
                    {
                        wb.Close();
                    }
                }
                if (path.Substring(path.Length - 4) == "xlsx")
                {
                    XSSFWorkbook wb = default;
                    try
                    {
                        using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
                        {
                            wb = new XSSFWorkbook(file);
                        }
                        ISheet sheet = wb.GetSheetAt(0);

                        for (int i = 16; i < 16 + 14; i++)
                        {
                            var cell = sheet.GetRow(i);
                            if(i%2 != 0)
                                Console.WriteLine($"{(i - 15)/2}. {cell.GetCell(21).ToString()}");
                        }
                        //{
                        //   
                        //}
                    }
                    finally
                    {
                        wb.Close();                        
                    }
                }
            }
            else { Console.WriteLine("Path is empty"); }
        }

        static void Main(string[] args)
        {
            GetFile(path: out path);
            GetTimetable(path: in path);
        }
    }
}