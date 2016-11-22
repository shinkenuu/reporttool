using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportTool.Business
{
    public interface IReport
    {
        void GenerateReport();

        void WriteToDisc();        
    }
}
