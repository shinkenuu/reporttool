using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportTool.Core
{
    public abstract class Api
    {


        protected string DecodeReport(string reportFormat, string reportName)
        {
            StringBuilder reportNameBuilder = new StringBuilder("ReportTool.Business.");
             
            switch (reportFormat.ToLower())
            {
                case "evolution":
                    Business.Evolution.EvolutionReportBuilder evolBuilder = new Business.Evolution.EvolutionReportBuilder();
                    reportNameBuilder.Append("Evolution.");
                    break;
                default:
                    throw new NotImplementedException();
            }
            
            switch (reportName.ToLower())
            {
                case "cvmonthly":
                    reportNameBuilder.Append("CV.MonthlyReport");
                    break;
                case "cvweekly":
                    reportNameBuilder.Append("CV.WeeklyReport");
                    break;
                case "msrphistory":
                    reportNameBuilder.Append("JATO.MsrpHistoryReport");
                    break;
                case "weightedmsrp":
                    reportNameBuilder.Append("JATO.WeightedMsrpReport");
                    break;
                default:
                    throw new NotImplementedException();
            }

            return reportNameBuilder.ToString();
        }







    }
}
