using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReportTool.Business;

namespace ReportTool.Core
{
    public class CliApi : Api
    {





        protected string[] Args { get; set; }





        public CliApi(string[] args)
        {
            if (args == null || args.Length < 1 || string.IsNullOrEmpty(args[0].Trim()))
            {
                Console.WriteLine("Invalid arguments");
                Environment.Exit(-1);
            }

            if (args[0] == "--help")
            {
                PrintHelp();
                Environment.Exit(0);
            }

            Args = args;
        }





        public void Api()
        {
            IReport report = null;
            Stopwatch watch = new Stopwatch();
            Business.Evolution.EvolutionReportBuilder evolBuilder = new Business.Evolution.EvolutionReportBuilder();
            string reportDecoded = DecodeReport(ExtractArgumentValue(Args, "reportFormat"), ExtractArgumentValue(Args, "reportName"));

            evolBuilder.ReportName = reportDecoded;
            evolBuilder.EarlyDate = DateTime.Parse(ExtractArgumentValue(Args, "earlyDate"));
            evolBuilder.LaterDate = DateTime.Parse(ExtractArgumentValue(Args, "laterDate"));
            Console.Write("Building " + reportDecoded);

            watch.Start();
            report = evolBuilder.Build();
            
            watch.Stop();

            if (report == null)
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
        private string ExtractArgumentValue(string[] args, string arg)
        {
            return args.FirstOrDefault(w => w.Contains("--" + arg)).Split(':')[1];
        }

        



        private void ConsoleWriteToRightSide(string msg)
        {
            Console.CursorLeft = Console.BufferWidth - msg.Length;
            Console.Write(msg);
        }





        private void PrintHelp()
        {
            Console.WriteLine("--reportFormat:[evolution]");
            Console.WriteLine("--reportName:[cvmonthly|cvweekly|msrphistory|weightedmsrp]");
            Console.WriteLine("--earlyDate:yyyyMMdd");
            Console.WriteLine("--laterDate:yyyyMMdd");
        }





    }
}
