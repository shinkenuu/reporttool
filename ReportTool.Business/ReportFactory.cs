using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReportTool.Repository;

namespace ReportTool.Business
{
    public class ReportFactory
    {
        public enum ReportType
        {
            CvWeekly,
            CvMonthly
        }

        protected DateTime EarlyDate { get;  set; }
        protected DateTime LaterDate { get; set; }
        
        public IReport Create(ReportType type)
        {
            switch(type)
            {
                case ReportType.CvMonthly:
                    CV.MonthlyReport cvMonthlyReport;
                    if (EarlyDate == null || LaterDate == null)
                    {
                        cvMonthlyReport = new CV.MonthlyReport();
                    }
                    else if(EarlyDate < LaterDate)
                    {
                        using (Repository.CV.CvEntity conn = new Repository.CV.CvEntity())
                        {
                            int earlyDate = EarlyDate.Year * 10000 + EarlyDate.Month * 100 + 1;
                            int laterDate = LaterDate.Year * 10000 + LaterDate.Month * 100 + 1;

                            var viewResults =  conn.MONTHLY_CV.Where(w => w.SAMPLE_DATE >= earlyDate && w.SAMPLE_DATE <= laterDate).OrderBy(v => v.MAKE).ThenBy(v => v.MODEL).ThenBy(v => v.VERSION).ThenBy(v => v.SAMPLE_DATE).ToList();

                            cvMonthlyReport = new CV.MonthlyReport(viewResults);
                        }                            
                    }
                    else
                    {
                        throw new InvalidOperationException("Can't create a report with given EarlyDate or LaterDate. EarlyDate:" + EarlyDate.ToString() + " LaterDate: " + LaterDate.ToString()); ;
                    }
                    return cvMonthlyReport as IReport;
                case ReportType.CvWeekly:
                    CV.WeeklyReport cvWeeklyReport;
                    IEnumerable<Repository.CV.WEEKLY_CV> cvWeeklyViewResults;
                    if (EarlyDate == null || LaterDate == null)
                    {
                        cvWeeklyReport = new CV.WeeklyReport();
                    }
                    else if (EarlyDate < LaterDate)
                    {
                        using (Repository.CV.CvEntity conn = new Repository.CV.CvEntity())
                        {
                            int earlyDate = EarlyDate.Year * 10000 + EarlyDate.Month * 100 + EarlyDate.Day;
                            int laterDate = LaterDate.Year * 10000 + LaterDate.Month * 100 + LaterDate.Day;

                            cvWeeklyViewResults = conn.WEEKLY_CV.Where(w => w.EXT_DATE >= earlyDate && w.EXT_DATE <= laterDate).OrderBy(v => v.MAKE).ThenBy(v => v.MODEL).ThenBy(v => v.VERSION).ThenBy(v => v.EXT_DATE).ToList();
                        }

                        cvWeeklyReport = new CV.WeeklyReport(cvWeeklyViewResults);
                    }
                    else
                    {
                        throw new InvalidOperationException("Can't create a report with given EarlyDate or LaterDate. EarlyDate:" + EarlyDate.ToString() + " LaterDate: " + LaterDate.ToString());
                    }
                    return cvWeeklyReport as IReport;
            }

            return null;
        }
        
        public void SetPeriod(DateTime earlyDate, DateTime laterDate)
        {
            if (earlyDate > laterDate)
            {
                throw new ArgumentOutOfRangeException("earlyDate can't be a date after laterDate");
            }

            EarlyDate = earlyDate;
            LaterDate = laterDate;
        }
        
    }
}
