using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JATO.Common.XlInterop
{
    public static class Utils
    {
        public static string XlCol(int arrayCol)
        {
            int div = arrayCol + 1;
            string colLetter = string.Empty;
            int mod = 0;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (int)((div - mod) / 26);
            }
            return colLetter;
        }

        public static int XlRow(int arrayRow)
        {
            return arrayRow + 1;
        }

    }
}
