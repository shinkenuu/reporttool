using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReportTool.Business;

namespace ReportTool.Core
{
    public class UiApi : Api
    {

        protected ReportCommand Command { get; set; }



        public UiApi(ReportCommand command)
        {
            if (command == null)
            {
                throw new ArgumentNullException();
            }

            else if (!command.IsReady)
            {
                throw new InvalidOperationException("Command is not ready yet");
            }

            Command = command;
        }


        public void Api()
        {
            IReport report = null;
            Business.Evolution.EvolutionReportBuilder evolBuilder = new Business.Evolution.EvolutionReportBuilder();
            string reportDecoded = DecodeReport(Command.ReportFormat, Command.ReportName);

            evolBuilder.ReportName = reportDecoded;
            evolBuilder.EarlyDate = Command.EarlyDate;
            evolBuilder.LaterDate = Command.LaterDate;
            Console.Write("Building " + reportDecoded);
            
            report = evolBuilder.Build();
            
            if (report == null)
            {
                return;
            }
            
            report.GenerateReport();
            
            report.WriteToDisc();
        }

        
    }

}
