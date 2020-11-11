using System;
using System.Collections.Generic;
using System.Net;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using System.IO;

namespace Timetable
{
    class Program
    {
        private HSSFWorkbook wb;
        private static string path = $"C:\\Users\\user\\Downloads\\Расписание на {GetDate()[0]} {months[GetDate()[1]]}";
        //private static List<string> formats = new List<string>() { ".xls", ".xlsx", ".pdf" };
        private static string domain;


        public static Dictionary<string, string> months = new Dictionary<string, string>()
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

        static string[] GetDate()
        {
            return DateTime.Today.ToString("d.MM.yyyy").Split('.');
        }

        static void GetHtml()
        {
            
            List<string> formats = new List<string>() { ".xls", ".xlsx", ".pdf" };
            foreach (string format in formats)
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create("http://www.mgkit.ru/studentu/raspisanie-zanatij/РАСПИСАНИЕ%20{GetDate()[0]}%20{months[GetDate()[1]]}%20{GetDate()[2]}{format}?attredirects=0&d=1");
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                if (response.StatusCode.ToString() == "OK")
                {
                    domain = domain = "http://www.mgkit.ru/studentu/raspisanie-zanatij/РАСПИСАНИЕ%20{GetDate()[0]}%20{months[GetDate()[1]]}%20{GetDate()[2]}{format}?attredirects=0&d=1";
                    WebClient wc = new WebClient();
                    wc.DownloadFile(domain, path + format);
                }
                else
                    Console.WriteLine(domain = "http://www.mgkit.ru/studentu/raspisanie-zanatij/РАСПИСАНИЕ%20{GetDate()[0]}%20{months[GetDate()[1]]}%20{GetDate()[2]}{format}?attredirects=0&d=1");
            }

        }

       

        static void Main(string[] args)
        {
            GetHtml();
        }
        
    }
}
