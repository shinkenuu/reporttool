using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XlUtils = JATO.Common.XlInterop.Utils;

namespace ReportTool.Business.Evolution
{
    public class ModelHeader
    {
        public string ModelSummaryFormula { get; set; }

        public int[,] ModelSummaryFormulaRanges { get; set; }


        public string MountModelSummaryFormula(int actualCol, int actualRow, int distinctSumOfVehiclesOfModel)
        {
            if (string.IsNullOrWhiteSpace(ModelSummaryFormula) || !ModelSummaryFormula.Contains("{0}"))
            {
                return null;
            }

            string mountedFormula = ModelSummaryFormula;

            for (int i = 0; i < ModelSummaryFormulaRanges.GetLength(0); i++)
            {
                mountedFormula = mountedFormula.Replace('{' + i.ToString() + '}', XlUtils.XlCol(actualCol + ModelSummaryFormulaRanges[i, 1]) + XlUtils.XlRow(actualRow + ModelSummaryFormulaRanges[i, 0]).ToString() 
                    + ':' + XlUtils.XlCol(actualCol + ModelSummaryFormulaRanges[i, 1]) + XlUtils.XlRow(actualRow + ModelSummaryFormulaRanges[i, 0] + distinctSumOfVehiclesOfModel - 1).ToString());
            }
            
            return mountedFormula;
        }

    }
}
