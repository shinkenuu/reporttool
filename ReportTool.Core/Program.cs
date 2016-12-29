using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReportTool.Business;
using System.Diagnostics;

namespace ReportTool.Core
{

    public class Program
    {
        static void Main(string[] args)
        {
            if (args == null || args.Length < 1)
            {
                return;
            }
            
            CliApi cli = new CliApi(args);
            cli.Api();
            return;
        }

        public static void Main(ReportCommand command)
        {
            UiApi ui = new UiApi(command);
            ui.Api();
            return;
        }
    }
}
