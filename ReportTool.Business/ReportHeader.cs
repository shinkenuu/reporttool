using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportTool.Business
{
    public abstract class ReportHeader
    {
        public enum SummaryOperation
        {
            Average,
            Sum
        }

        public string HeaderName { get; set; }        
        public string NumberFormat { get; set; }
        public byte Offset { get; set; }
    }
}
