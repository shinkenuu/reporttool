using Microsoft.VisualStudio.TestTools.UnitTesting;
using JatoDbConverter = JATO.Common.Database.Converter;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JATO.Common.Database.Tests
{
    [TestClass()]
    public class ConverterTests
    {
        [TestMethod()]
        public void FromIntToDateTest()
        {
            //Declare
            List<int> intList = new List<int>();
            List<DateTime> dateList = new List<DateTime>();

            //Arrange
            intList.Add(20150131);
            intList.Add(20160228);
            intList.Add(20021201);
            intList.Add(19970405);

            dateList.Add(new DateTime(2015, 01, 31));
            dateList.Add(new DateTime(2016, 02, 28));
            dateList.Add(new DateTime(2002, 12, 01));
            dateList.Add(new DateTime(1997, 04, 05));

            //Assert
            for (int i = 0; i < intList.Count; i++)
            {
                Assert.AreEqual(dateList[0], JatoDbConverter.FromIntToDate(intList[0]));
            }
        }
    }
}