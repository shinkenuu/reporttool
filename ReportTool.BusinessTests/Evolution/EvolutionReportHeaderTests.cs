using Microsoft.VisualStudio.TestTools.UnitTesting;
using ReportTool.Business.Evolution;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportTool.Business.Evolution.Tests
{
    [TestClass()]
    public class EvolutionReportHeaderTests
    {
        [TestMethod()]
        public void MountModelSummaryFormulaTest()
        {
            //Declare
            //EvolutionReportHeader header1           = new EvolutionReportHeader("header1", "#,###", 0, "=IF(ISERROR(SUMPRODUCT({0},{1})/SUM({1})), \" - \",SUMPRODUCT({0},{1})/SUM({1}))");
            //EvolutionReportHeader header2           = new EvolutionReportHeader("header2", "#,###", 1, "=IF(ISERROR(SUM({0})), \" - \",SUM({0}))");
            //EvolutionReportHeader messedHeader1  = new EvolutionReportHeader("incorrectHeader1", "#,###", 2, "=IF(ISERROR(SUMPRODUCT({0},1)/SUM({1)), \" - \",SUMPRODUCT({0},{3})/SUM({2}))");
            //EvolutionReportHeader messedHeader2  = new EvolutionReportHeader("incorrectHeader2", "#,###", 3, "=IF(ISERROR(SUM({0})), \" - \",SUM({2}))");

            ////Arrange
            //header1.FormulaParameters = new string[] { "param0", "param1" };
            //header2.FormulaParameters = new string[] { "param0" };
            //messedHeader1.FormulaParameters = new string[] { "param0", "param1", "param2", "param3" };
            //messedHeader2.FormulaParameters = new string[] { "param0", "param1", "param2" };

            ////Assert
            //Assert.AreEqual("=IF(ISERROR(SUMPRODUCT(param0,param1)/SUM(param1)), \" - \",SUMPRODUCT(param0,param1)/SUM(param1))", header1.MountModelSummaryFormula(header1.ModelSummaryFormula, header1.FormulaParameters));
            //Assert.AreEqual("=IF(ISERROR(SUM(param0)), \" - \",SUM(param0))", header2.MountModelSummaryFormula(header2.ModelSummaryFormula, header2.FormulaParameters));
            //Assert.AreEqual("=IF(ISERROR(SUMPRODUCT(param0,1)/SUM({1)), \" - \",SUMPRODUCT(param0,param3)/SUM(param2))", messedHeader1.MountModelSummaryFormula(messedHeader1.ModelSummaryFormula, messedHeader1.FormulaParameters));
            //Assert.AreEqual("=IF(ISERROR(SUM(param0)), \" - \",SUM(param2))", messedHeader2.MountModelSummaryFormula(messedHeader2.ModelSummaryFormula, messedHeader2.FormulaParameters));
        }






        [TestMethod()]
        public void GetAmountOfParametersInFormulaTest()
        {
        //    //Declare
        //    //Have a correct formula
        //    EvolutionReportHeader header = new EvolutionReportHeader("test", "#,###", 0, "=IF(ISERROR(SUMPRODUCT({0},{1})/SUM({1})), \" - \",SUMPRODUCT({0},{1})/SUM({1}))");

        //    //Have valid {0} and {2} params, but none {1}
        //    EvolutionReportHeader incorrectHeader = new EvolutionReportHeader("test", "#,###", 0, "=IF(ISERROR(SUMPRODUCT(0},{1)/SUM({2})), \" - \",SUMPRODUCT({0},1)/SUM(1))");

        //    //Arrange

        //    //Assert
        //    Assert.AreEqual(2, header.GetAmountOfParametersInFormula(header.ModelSummaryFormula));
        //    Assert.AreEqual(1, incorrectHeader.GetAmountOfParametersInFormula(incorrectHeader.ModelSummaryFormula));
        }
    }
}