using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportTool.Core
{
    public class ReportCommand
    {
        public enum ReportNames
        {
            CvMonthly,
            CvWeekly,
            MsrpHistory,
            WeightedMsrp
        }

        public enum ReportFormats
        {
            Evolution
        }


        public string ReportName { get; set; }
        public string ReportFormat { get; set; }


        private DateTime _earlyDate = DateTime.Today;
        private DateTime _laterDate = DateTime.Today;

        public DateTime EarlyDate
        {
            get
            {
                return _earlyDate;
            }

            set
            {
                if(_laterDate > value)
                {
                    _earlyDate = value;
                }
            }
        }


        public DateTime LaterDate
        {
            get
            {
                return _laterDate;
            }

            set
            {
                if (_earlyDate < value)
                {
                    _laterDate = value;
                }
            }
        }


        public bool IsReady
        {
            get
            {
                return ReportName != null && ReportFormat != null;
            }
        }
        

        public string[] ToMainArgs()
        {
            if(!IsReady)
            {
                throw new InvalidOperationException("Command is not ready yet");
            }

            List<string> argsList = new List<string>();

            argsList.Add("--reportFormat:" + ReportFormat);
            argsList.Add("--reportName:" + ReportName);

            if(ReportFormat == "evolution")
            {
                argsList.Add("--earlyDate:" + EarlyDate.ToString("yyyy-MM-dd"));
                argsList.Add("--laterDate:" + LaterDate.ToString("yyyy-MM-dd"));
            }

            return argsList.ToArray();
        }
    }
}
