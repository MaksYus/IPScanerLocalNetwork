
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net.Http.Headers;
using System.IO;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using NUnit.Framework;
using System.Net.NetworkInformation;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Drawing;
using System.Net;

namespace Testing_1
{



    class Program
    {
        static void PingCompletedCallback(object sender, PingCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                Console.WriteLine("Ping canceled.");
                ((AutoResetEvent)e.UserState).Set();
            }

            if (e.Error != null)
            {
                Console.WriteLine("Ping failed:");
                Console.WriteLine(e.Error.ToString());

                ((AutoResetEvent)e.UserState).Set();
            }


            PingReply reply = e.Reply;

            DisplayReply(reply);

            ((AutoResetEvent)e.UserState).Set();

        }

        static void DisplayReply(PingReply reply)
        {
            if (reply == null)
                return;
            Adress adress = new Adress();
            adress.IP = Ip;
            adress.Status = reply.Status.ToString();
            adress.HostName = NSLookup(adress.IP);
            adresses.Add(adress);
        }
        public struct Adress
        {
            public string IP;
            public string Status;
            public string HostName;
        }
        static object lockWrite = new object();
        public static Queue<string> ip = new Queue<string>();
        public static string Ip
        {
            get
            {
                lock (lockWrite)
                {
                   
                        return ip.Dequeue();
                    
                }
            }
            
            set
            {
                lock (lockWrite)
                {
                    ip.Enqueue(value);
                }
            }
        }

        public static List<Adress> adresses = new List<Adress>();

        static void F(string[] args)
        {
            if (args.Length == 0)
                throw new ArgumentException("Ping needs a host or IP Address.");

            string who = args[0];
            AutoResetEvent waiter = new AutoResetEvent(false);

            Ping pingSender = new Ping();

            pingSender.PingCompleted += new PingCompletedEventHandler(PingCompletedCallback);

            string data = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";
            byte[] buffer = Encoding.ASCII.GetBytes(data);

            int timeout = 1000;
            PingOptions options = new PingOptions(64, true);
            pingSender.SendAsync(who, timeout, buffer, options, waiter);
            

            waiter.WaitOne();

        }
        static string NSLookup(string args)
        {
            try
            {
                IPHostEntry ipEntry;
                IPAddress[] ipAddr;
                char[] alpha = "aAbBcCdDeEfFgGhHiIjJkKlLmMnNoOpPqQrRsStTuUvVwWxXyYzZ-".ToCharArray();
                if (args.IndexOfAny(alpha) != -1)
                {
                    ipEntry = Dns.GetHostByName(args);
                    ipAddr = ipEntry.AddressList;
                    int i = 0;
                    int len = ipAddr.Length;
                    return args;
                }
                else
                {
                    ipEntry = Dns.Resolve(args);
                    return ipEntry.HostName;
                }
            }
            catch
            {
                return "";
            }
        }

        static void Generate(object obj)
        {
            Param par = (Param)obj;
            for (int i =0; i<par.IPs.Count; i++)
            {
                if (0 == i) Console.WriteLine("Поток №" + par.number + " Стартовал, кол-во задач: " + par.IPs.Count);
                if (par.IPs.Count / 4 == i) Console.WriteLine("Поток №" + par.number + " Выполнен на 25%");
                if (par.IPs.Count / 2 == i ) Console.WriteLine("Поток №" + par.number + " Выполнен на 50%");
                if (par.IPs.Count * 3 / 4 == i) Console.WriteLine("Поток №" + par.number + " Выполнен на 75%");
                Ip = par.IPs[i];
                F(new string[] { par.IPs[i] });
                if (par.IPs.Count - 1  == i) Console.WriteLine("Поток №" + par.number + " Завершил работу");
            }
        }


        struct Param
        {
            public List<string> IPs;
            public int number;
        }
        static void Main(string[] args)
        {
            string path = Environment.CurrentDirectory;
            Excel.Application exData = new Excel.Application();
            exData.Workbooks.Open(path + @"\Adresses.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);
            exData.DisplayAlerts = false;
            Excel.Worksheet sheet = (Excel.Worksheet)exData.Worksheets.get_Item(1);

            string Host = Dns.GetHostName();
            string IP = Dns.GetHostByName(Host).AddressList[0].ToString();

            string LocalNetwork = IP.Remove(IP.LastIndexOf('.')+1,IP.Length - IP.LastIndexOf('.') - 1);

            List<Thread> Threads = new List<Thread>();
            int countThread = 10;
            List<List<string>> IPList = new List<List<string>>();
            for (int i = 0; i < countThread; i++)
            {
                List<string> ips = new List<string>();
                for (int j = 0; j < 255 / countThread; j++)
                {
                    string ipstr = LocalNetwork + (j + 255 / countThread * i);
                    ips.Add(ipstr);
                }
                IPList.Add(ips);
            }
            

            for (int i = 0; i < countThread; i++)
            {
                Thread myThread = new Thread(new ParameterizedThreadStart(Generate));
                myThread.Start(new Param { IPs = IPList[i] , number = i});
                Threads.Add(myThread);
            }

            foreach (var item in Threads)
            {
                if (item.IsAlive) item.Join();
            }
            Console.WriteLine("Все потоки прекратили работу");
            List<string> ojerj = new List<string>();
            for (int i = ((255 / countThread - 1) + 255 / countThread * (countThread - 1))+1; i < 256; i++)
            {
                ojerj.Add(LocalNetwork + (i));
            }
            Generate(new Param { IPs = ojerj,number = -1 });
            Console.WriteLine("Пинг закончился");




            Console.WriteLine("Заполнение таблицы");
            for (int i = 0; i < adresses.Count; i++)
            {
                sheet.Cells[i + 2, 1] = adresses[i].IP;
                sheet.Cells[i + 2, 2] = adresses[i].Status;
                if (adresses[i].Status == "Success")
                {
                    sheet.Cells[i + 2, 2].Interior.Color = ColorTranslator.ToOle(Color.FromArgb(0, 230, 0));
                }
                else
                {
                    sheet.Cells[i + 2, 2].Interior.Color = ColorTranslator.ToOle(Color.FromArgb(230, 0, 0));
                }
                sheet.Cells[i + 2, 3] = adresses[i].HostName;
            }

            exData.Application.ActiveWorkbook.SaveAs(path + @"\Adresses.xlsx", Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            exData.Quit();
            Console.WriteLine("Таблица заполнена\n\nВыполнение программы завершено!");
            Console.ReadKey();
        }
    }
}
