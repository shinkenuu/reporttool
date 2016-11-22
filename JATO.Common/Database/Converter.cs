using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JATO.Common.Database
{
    public static class Converter
    {
        public static DateTime FromIntToDate(int date)
        {
            if (date < 19000101)
            {
                return DateTime.MinValue;
            }

            //int year = date / 10000;
            //int month = (date % 10000) / 100;
            //int day = date % 100;

            return new DateTime(date / 10000, (date % 10000) / 100, date % 100);
        }


        public static IEnumerable<DateTime> FromIntToDate(IEnumerable<int> dates)
        {
            List<DateTime> list = new List<DateTime>();

            foreach (int date in dates)
            {
                if (date < 19000101)
                {
                    return null;
                }

                list.Add(new DateTime(date / 10000, (date % 10000) / 100, date % 100));
            }

            return list;
        }
    }
}
