using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportTool.Business.Evolution
{

    public class EvolutionReportBuilder : IReportBuilder
    {

        private string _reportName;
        private DateTime? _earlyDate;
        private DateTime? _laterDate;
        
        public DateTime EarlyDate
        {
            get
            {
                return _earlyDate ?? DateTime.MinValue;
            }

            set
            {
                if (LaterDate != null && value > LaterDate)
                {
                    _earlyDate = null;
                    throw new InvalidOperationException("EarlyDate cant be bigger than LaterDate");
                }

                _earlyDate = value;
            }
        }

        public DateTime LaterDate
        {
            get
            {
                return _laterDate ?? DateTime.MaxValue;
            }

            set
            {

                if (EarlyDate != null && value < EarlyDate)
                {
                    _laterDate = null;
                    throw new InvalidOperationException("LaterDate cant be lower than EarlyDate");
                }

                _laterDate = value;
            }
        }

        public string ReportName
        {
            get
            {
                return _reportName;
            }

            set
            {
                switch (value)
                {
                    case "ReportTool.Business.Evolution.CV.MonthlyReport":
                    case "ReportTool.Business.Evolution.CV.WeeklyReport":
                    case "ReportTool.Business.Evolution.JATO.MsrpHistoryReport":
                    case "ReportTool.Business.Evolution.JATO.WeightedMsrpReport":
                        _reportName = value;
                        break;
                    default: 
                        _reportName = null;
                        throw new NotImplementedException(value);
                }
            }
        }





        public IReport Build()
        {
            if (EarlyDate == DateTime.MinValue && LaterDate == DateTime.MaxValue)
            {
                throw new InvalidOperationException("Can't build a report with given EarlyDate or LaterDate. EarlyDate:" + EarlyDate.ToString() + " LaterDate: " + LaterDate.ToString());
            }
            else if(ReportName == null)
            {
                throw new InvalidOperationException("Can't build a unspecified report");
            }

            switch (ReportName)
            {
                case "ReportTool.Business.Evolution.CV.MonthlyReport":
                    CV.MonthlyReport cvMonthlyReport;
                    cvMonthlyReport = new CV.MonthlyReport();
                    IEnumerable<Repository.CV.MONTHLY_CV> cvMonthlyViewResults;
                    using (Repository.CV.CvEntity conn = new Repository.CV.CvEntity())
                    {
                        int earlyDate = EarlyDate.Year * 10000 + EarlyDate.Month * 100 + 1;
                        int laterDate = LaterDate.Year * 10000 + LaterDate.Month * 100 + 1;

                        cvMonthlyViewResults = conn.MONTHLY_CV.Where(w => w.SAMPLE_DATE >= earlyDate && w.SAMPLE_DATE <= laterDate).OrderBy(v => v.MAKE).ThenBy(v => v.MODEL).ThenBy(v => v.VERSION).ThenBy(v => v.SAMPLE_DATE).ToList();
                    }
                    cvMonthlyReport = new CV.MonthlyReport(cvMonthlyViewResults);
                    return cvMonthlyReport as IReport;
                case "ReportTool.Business.Evolution.CV.WeeklyReport":
                    CV.WeeklyReport cvWeeklyReport;
                    IEnumerable<Repository.CV.WEEKLY_CV> cvWeeklyViewResults;
                    using (Repository.CV.CvEntity conn = new Repository.CV.CvEntity())
                    {
                        int earlyDate = EarlyDate.Year * 10000 + EarlyDate.Month * 100 + EarlyDate.Day;
                        int laterDate = LaterDate.Year * 10000 + LaterDate.Month * 100 + LaterDate.Day;

                        cvWeeklyViewResults = conn.WEEKLY_CV.Where(w => w.EXT_DATE >= earlyDate && w.EXT_DATE <= laterDate).OrderBy(v => v.MAKE).ThenBy(v => v.MODEL).ThenBy(v => v.VERSION).ThenBy(v => v.EXT_DATE).ToList();
                    }
                    cvWeeklyReport = new CV.WeeklyReport(cvWeeklyViewResults);
                    return cvWeeklyReport as IReport;
                case "ReportTool.Business.Evolution.JATO.MsrpHistoryReport":
                    JATO.MsrpHistoryReport msrpHistoryReport;
                    IEnumerable<Repository.RT.MSRP_HISTORY> msrpHistorViewResults;
                    using (Repository.RT.RtEntity conn = new Repository.RT.RtEntity())
                    {
                        int earlyDate = EarlyDate.Year * 10000 + EarlyDate.Month * 100 + 1;
                        int laterDate = LaterDate.Year * 10000 + LaterDate.Month * 100 + 1;

                        msrpHistorViewResults = conn.MSRP_HISTORY.Where(w => w.DATADATE >= earlyDate && w.DATADATE <= laterDate).OrderBy(v => v.MAKE).ThenBy(v => v.MODEL).ThenBy(v => v.VERSION).ThenBy(v => v.DATADATE).ToList();
                    }
                    msrpHistoryReport = new JATO.MsrpHistoryReport(msrpHistorViewResults);
                    return msrpHistoryReport as IReport;
                case "ReportTool.Business.Evolution.JATO.WeightedMsrpReport":
                    JATO.WeightedMsrpReport weightedMsrpReport;
                    IEnumerable<Repository.RT.MONTHLY_MSRP> monthlyMsrpViewResults;
                    using (Repository.RT.RtEntity conn = new Repository.RT.RtEntity())
                    {
                        int earlyDate = EarlyDate.Year * 10000 + EarlyDate.Month * 100 + 1;
                        int laterDate = LaterDate.Year * 10000 + LaterDate.Month * 100 + 1;

                        monthlyMsrpViewResults = conn.MONTHLY_MSRP.Where(w => w.SAMPLE_DATE >= earlyDate && w.SAMPLE_DATE <= laterDate).OrderBy(v => v.MAKE).ThenBy(v => v.MODEL).ThenBy(v => v.VERSION).ThenBy(v => v.SAMPLE_DATE).ToList();
                    }
                    weightedMsrpReport = new JATO.WeightedMsrpReport(monthlyMsrpViewResults);
                    return weightedMsrpReport as IReport;
            }

            return null;
        }

    }
}
