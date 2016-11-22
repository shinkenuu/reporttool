using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReportTool.UI;
using ReportTool.Business;
using System.Diagnostics;

namespace ReportTool.Core
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args == null || args.Length < 1)
            {
                Console.WriteLine("Invalid arguments");
                return;
            }

            ReportFactory factory = new ReportFactory();
            IReport report = null;
            Stopwatch watch = new Stopwatch();
            
            switch (ExtractArgumentValue(args, "reportFormat").ToLower())
            {
                case "cvmonthly":
                    factory.SetPeriod(DateTime.Parse(ExtractArgumentValue(args, "earlyDate")), DateTime.Parse(ExtractArgumentValue(args, "laterDate")));
                    Console.Write("Creating CV - Monthly Report");
                    watch.Start();
                    report = factory.Create(ReportFactory.ReportType.CvMonthly);
                    break;
                case "cvweekly":
                    factory.SetPeriod(DateTime.Parse(ExtractArgumentValue(args, "earlyDate")), DateTime.Parse(ExtractArgumentValue(args, "laterDate")));
                    Console.Write("Creating CV - Weekly Report");
                    watch.Start();
                    report = factory.Create(ReportFactory.ReportType.CvWeekly);
                    break;
            }

            watch.Stop();

            if(report == null)
            {
                return;
            }

            ConsoleWriteToRightSide(string.Format("[{0:00}:{1:00}.{2:00}]", watch.Elapsed.Minutes, watch.Elapsed.Seconds, watch.Elapsed.Milliseconds));

            watch.Reset();
            Console.Write("Generating Report");
            watch.Start();
            report.GenerateReport();
            watch.Stop();

            ConsoleWriteToRightSide(string.Format("[{0:00}:{1:00}.{2:00}]", watch.Elapsed.Minutes, watch.Elapsed.Seconds, watch.Elapsed.Milliseconds));

            Console.Write("Writing Report to disk");
            report.WriteToDisc();
        }

        /// <summary>
        /// Extracts a argument value from the arguments array
        /// </summary>
        /// <param name="args">The array of arguments</param>
        /// <param name="arg">The argument to look for ot</param>
        /// <returns>The value of the argument specified, if any</returns>
        private static string ExtractArgumentValue(string[] args, string arg)
        {
            return args.FirstOrDefault(w => w.Contains("--" + arg)).Split(':')[1];
        }

        private static void ConsoleWriteToRightSide(string msg)
        {
            Console.CursorLeft = Console.BufferWidth - msg.Length;
            Console.Write(msg);
        }
    }
}
