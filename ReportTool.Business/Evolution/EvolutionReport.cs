using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using XlUtils = JATO.Common.XlInterop.Utils;

namespace ReportTool.Business.Evolution
{
    public abstract class EvolutionReport : Report, IReport
    {
        
        protected enum constPostns : int
        {
            TitleHeaderRow = 0,
            TimeHeaderRow = 1,
            InfoHeaderRow = 2,
            FirstSampleRow = 3,
            VehicleCol = 0,
            FirstSampleCol = 1,
            // Summary Sheet Constants
            summMakeCol = 0,
            summTimeHeaderRow = 0
        };


        /// <summary>
        /// Comes from DB as INT. Format: yyyyMMdd
        /// </summary>
        protected IEnumerable<DateTime> SampleDates;
        
        protected List<EvolutionReportHeader> SampleHeaders;

        /// <summary>
        /// Mark-Up Column to sign that this row contains vehicle description
        /// </summary>
        protected int VehDescRowMarkUpColumn = -1;
        






        /// <summary>
        /// Calls GenerateEvolutionReport()
        /// </summary>
        public override void GenerateReport()
        {
            GenerateEvolutionReport();
        }

        public abstract void GenerateEvolutionReport();




        protected void SetupNewWorksheet(string newMakeName, int sheetIndex, int matrixWidth, int matrixHeight)
        {
            Matrix = null;
            Matrix = new string[matrixHeight, matrixWidth];

            //Initial settings for the new worksheet
            XlWorksheet = (Excel.Worksheet)XlSheets[sheetIndex];
            XlWorksheet.UsedRange.Clear();
            XlWorksheet.Name = newMakeName;
        }





        #region Headers

        protected void WriteWorksheetHeaders(string worksheetMakeName, string reportName, string creationDateFormat, string sampleDateFormat)
        {
            Matrix[(int)constPostns.TitleHeaderRow, (int)constPostns.VehicleCol] = "Data from " + SampleDates.First().ToString(creationDateFormat) + " to " + SampleDates.Last().ToString(creationDateFormat);
            Matrix[(int)constPostns.TimeHeaderRow, (int)constPostns.VehicleCol] = "Vehicles";

            Matrix[(int)constPostns.TitleHeaderRow, (int)constPostns.FirstSampleCol] = worksheetMakeName + " - " + reportName;

            int col = (int)constPostns.FirstSampleCol;

            foreach (DateTime sampleDate in SampleDates)
            {
                Matrix[(int)constPostns.TimeHeaderRow, col] = sampleDate.Date.ToString(sampleDateFormat);

                foreach (string headerName in SampleHeaders.Select(h => h.HeaderName).Distinct())
                {
                    Matrix[(int)constPostns.InfoHeaderRow, col++] = headerName;
                }
            }

            //insert JATO brand image
        }


        protected void WriteModelHeader(string modelName, int distinctSumOfVehiclesOfModel, int modelHeaderRow)
        {
            int firstColumnOfSample;
            int absoluteColumnOfHeader;

            Matrix[modelHeaderRow, (int)constPostns.VehicleCol] = modelName;

            for (int sampleDateIdx = 0; sampleDateIdx < SampleDates.Count(); sampleDateIdx++)
            {
                firstColumnOfSample = (int)constPostns.FirstSampleCol + sampleDateIdx * SampleHeaders.Count;

                foreach (EvolutionReportHeader header in SampleHeaders)
                {
                    absoluteColumnOfHeader = firstColumnOfSample + header.Offset;
                    Matrix[modelHeaderRow, absoluteColumnOfHeader] = header.ModelHeader.MountModelSummaryFormula(absoluteColumnOfHeader, modelHeaderRow, distinctSumOfVehiclesOfModel);
                }
            }
        }

        #endregion




        #region End Work

        protected void FinalizeActiveWorksheet(int lastRowIndex)
        {
            FillEmptyVehicleCells(lastRowIndex);
            WriteMatrixToExcel();
            AddHeaderExcelFeatures(XlUtils.XlRow(lastRowIndex));
            AddSampleExcelFeatures(XlUtils.XlRow(lastRowIndex));
            AddEvolutionReportExcelFeatures();
        }


        protected void FillEmptyVehicleCells(int versionRow)
        {
            byte sampleDateIdx = 0;

            foreach (DateTime sampleDate in SampleDates)
            {
                if (Matrix[versionRow, (int)constPostns.FirstSampleCol + SampleHeaders.Count * sampleDateIdx] == null)
                {
                    foreach (EvolutionReportHeader header in SampleHeaders)
                    {
                        Matrix[versionRow, (int)constPostns.FirstSampleCol + SampleHeaders.Count * sampleDateIdx + header.Offset] = "?";
                    }
                }

                sampleDateIdx++;
            }
        }

        #endregion





        protected void AddHeaderExcelFeatures(int xlRowsAmount)
        {
            Excel.Range rang;
            
            //All headers
            rang = XlWorksheet.get_Range(XlUtils.XlCol((int)constPostns.VehicleCol) + XlUtils.XlRow((int)constPostns.TitleHeaderRow) + ":" + XlUtils.XlCol(VehDescRowMarkUpColumn) + XlUtils.XlRow((int)constPostns.InfoHeaderRow));
            rang.Font.Name = "Arial";
            rang.Font.Size = "12";
            rang.Font.Bold = true;

            //Time and Info Headers
            rang = XlWorksheet.get_Range(XlUtils.XlCol((int)constPostns.VehicleCol) + XlUtils.XlRow((int)constPostns.TimeHeaderRow) + ":" + XlUtils.XlCol(VehDescRowMarkUpColumn) + XlUtils.XlRow((int)constPostns.InfoHeaderRow));
            rang.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = 3d;
            rang.Interior.Color = Color.FromArgb(147, 0, 4); // dark red
            rang.Font.Color = Color.FromArgb(241, 242, 242); // white
            
            // Report Title
            rang = XlWorksheet.get_Range(XlUtils.XlCol((int)constPostns.FirstSampleCol) + XlUtils.XlRow((int)constPostns.TitleHeaderRow) + ":" + XlUtils.XlCol((int)constPostns.FirstSampleCol + SampleDates.Count() * SampleHeaders.Count) + XlUtils.XlRow((int)constPostns.TitleHeaderRow));
            rang.Merge();
            rang.Font.Size = "28";
            
            //"Vehicles" merge between time header row and info header row
            rang = XlWorksheet.get_Range(XlUtils.XlCol((int)constPostns.VehicleCol) + XlUtils.XlRow((int)constPostns.TimeHeaderRow) + ":" + XlUtils.XlCol((int)constPostns.VehicleCol) + XlUtils.XlRow((int)constPostns.InfoHeaderRow));
            rang.Merge();

            int maxSampleHeaderOffset = SampleHeaders.Max(h => h.Offset);

            for (int sampleDateIdx = 0; sampleDateIdx < SampleDates.Count(); sampleDateIdx++)
            {
                //Merge Time Header cells of SampleDates
                rang = XlWorksheet.get_Range(XlUtils.XlCol((int)constPostns.FirstSampleCol + sampleDateIdx * SampleHeaders.Count) + XlUtils.XlRow((int)constPostns.TimeHeaderRow) + ":" + XlUtils.XlCol((int)constPostns.FirstSampleCol + sampleDateIdx * SampleHeaders.Count + maxSampleHeaderOffset) + XlUtils.XlRow((int)constPostns.TimeHeaderRow));
                rang.Merge();

                //Paint InfoHeader and TimeHeader left border white
                rang = XlWorksheet.get_Range(XlUtils.XlCol((int)constPostns.FirstSampleCol + sampleDateIdx * SampleHeaders.Count) + XlUtils.XlRow((int)constPostns.TimeHeaderRow) + ":" + XlUtils.XlCol((int)constPostns.FirstSampleCol + sampleDateIdx * SampleHeaders.Count + maxSampleHeaderOffset) + XlUtils.XlRow((int)constPostns.InfoHeaderRow));
                rang.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = 3d;
                rang.Borders[Excel.XlBordersIndex.xlEdgeLeft].Color = Color.FromArgb(241, 242, 242); // white
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="xlRowsAmount"></param>
        /// <param name="xlRightMostColumn"></param>
        /// <param name="headersNumberFormats">#,### for numbers and #.00% for percentages. Index must match the SampleHeaders index. All with be set as #,### if left null or empty</param>
        protected void AddSampleExcelFeatures(int xlRowsAmount)
        {
            Excel.Range rang;
            byte sampleDateIdx = 0;
            
            foreach (var sampleDate in SampleDates)
            {
                string curXlColumnLetter = null;

                foreach (EvolutionReportHeader header in SampleHeaders)
                {
                    curXlColumnLetter = XlUtils.XlCol((int)constPostns.FirstSampleCol + SampleHeaders.Count * sampleDateIdx + header.Offset);
                    rang = XlWorksheet.get_Range((curXlColumnLetter + XlUtils.XlRow((int)constPostns.FirstSampleRow) + ":" + curXlColumnLetter + xlRowsAmount));
                    rang.NumberFormat = header.NumberFormat;
                }

                sampleDateIdx++;
            }

            for (int row = (int)constPostns.FirstSampleRow; row < xlRowsAmount; row++)
            {
                //If not a vehicleDescription row
                if (Matrix[row, VehDescRowMarkUpColumn] != "x")
                {
                    rang = XlWorksheet.get_Range(XlUtils.XlCol((int)constPostns.VehicleCol) + XlUtils.XlRow(row) + ":" + XlUtils.XlCol(VehDescRowMarkUpColumn) + XlUtils.XlRow(row));
                    rang.Interior.Color = Color.FromArgb(72, 73, 78); //dark gray
                    rang.Font.Color = Color.FromArgb(241, 242, 242); // white
                    rang.Font.Bold = true;
                    rang.Font.Size = "12";
                }

                //alternate rows to make the background color grid
                else if (row % 2 != 0)
                {
                    rang = XlWorksheet.get_Range(XlUtils.XlCol((int)constPostns.VehicleCol) + XlUtils.XlRow(row) + ":" + XlUtils.XlCol(VehDescRowMarkUpColumn) + XlUtils.XlRow(row));
                    rang.Interior.Color = Color.FromArgb(220, 221, 222); //light gray
                }
            }
        }

        /// <summary>
        /// Aligns, center, auto-fit and split cells
        /// </summary>
        /// <param name="xlRowsAmount"></param>
        /// <param name="xlRightMostColumn"></param>
        protected void AddEvolutionReportExcelFeatures()
        {
            Excel.Range rang;
            
            rang = XlWorksheet.UsedRange;

            rang.EntireColumn.AutoFit();
            rang.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            rang.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            rang.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = 3d;
            rang.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 3d;
            XlWorksheet.get_Range(XlUtils.XlCol((int)constPostns.VehicleCol) + XlUtils.XlRow((int)constPostns.FirstSampleRow) + ":" + XlUtils.XlCol((int)constPostns.VehicleCol) + XlUtils.XlRow(Matrix.GetLength(0))).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            XlWorksheet.Cells[1, XlUtils.XlRow(VehDescRowMarkUpColumn)].EntireColumn.Hidden = true;
            XlWorksheet.Activate();
            XlWorksheet.Application.ActiveWindow.SplitRow = XlUtils.XlRow((int)constPostns.InfoHeaderRow);
            XlWorksheet.Application.ActiveWindow.SplitColumn = XlUtils.XlRow((int)constPostns.VehicleCol);
            XlWorksheet.Application.ActiveWindow.FreezePanes = true;
        }

    }
}
