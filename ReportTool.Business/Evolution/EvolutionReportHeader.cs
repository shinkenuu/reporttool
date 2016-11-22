using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportTool.Business.Evolution
{
    public class EvolutionReportHeader : ReportHeader
    {
        public ModelHeader ModelHeader { get; set; }
        
        public EvolutionReportHeader(string headerName, string numberFormat, byte offset)
        {
            HeaderName = headerName;
            NumberFormat = numberFormat;
            Offset = offset;
            ModelHeader = new ModelHeader();
        }
        
    }

}
