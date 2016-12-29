using Microsoft.VisualStudio.TestTools.UnitTesting;
using ReportTool.Business.CV;
using ReportTool.Repository;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;



namespace ReportTool.Business.CV.Tests
{
    [TestClass()]
    public class WeeklyReportTests
    {
        [TestMethod()]
        public void WeeklyReportTest()
        {
            //Declare
            List<Repository.CV.WEEKLY_CV> entityList = new List<Repository.CV.WEEKLY_CV>();
            Repository.CV.WEEKLY_CV weeklyEntityExample = new Repository.CV.WEEKLY_CV();
            weeklyEntityExample.MAKE = "Fake Make";
            weeklyEntityExample.MODEL = "Fake Model";
            weeklyEntityExample.VERSION = "Fake Version";
            weeklyEntityExample.PY = "2016";
            weeklyEntityExample.MY = "2017";
            weeklyEntityExample.BODY_TYPE = "Fake Body Type";
            weeklyEntityExample.MSRP_PLUS_OPC = 40000;
            weeklyEntityExample.TP = 50000;
            weeklyEntityExample.EXT_DATE = 20160501;

            //Arrange
            for (byte i = 0; i < 29; i++)
            {
                Repository.CV.WEEKLY_CV weeklyEntity = new Repository.CV.WEEKLY_CV();

                weeklyEntity.MAKE = weeklyEntityExample.MAKE + i.ToString();
                weeklyEntity.MODEL = weeklyEntityExample.MODEL + i.ToString();
                weeklyEntity.VERSION = weeklyEntityExample.VERSION + i.ToString();
                weeklyEntity.PY = weeklyEntityExample.PY + i.ToString();
                weeklyEntity.MY = weeklyEntityExample.MY + i.ToString();
                weeklyEntity.BODY_TYPE = weeklyEntityExample.BODY_TYPE + i.ToString();
                weeklyEntity.MSRP_PLUS_OPC = weeklyEntityExample.MSRP_PLUS_OPC = +i;
                weeklyEntity.TP = weeklyEntityExample.TP = +i;
                weeklyEntity.EXT_DATE = weeklyEntityExample.EXT_DATE = +i;

                entityList.Add(weeklyEntity);
            }

            //WeeklyReport weeklyReport = new WeeklyReport(entityList);
            //weeklyReport.GenerateReport();

            //Assert
        
        }

        [TestMethod()]
        public void WeeklyReportTest1()
        {
            //Assert.Fail();
        }

        [TestMethod()]
        public void GenerateReportTest()
        {
            //Assert.Fail();
        }

        [TestMethod()]
        public void SetDataTest()
        {
            //Assert.Fail();
        }
    }
}