using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using XlUtils = JATO.Common.XlInterop.Utils;
using JATODbConverter = JATO.Common.Database.Converter;

namespace ReportTool.Business.CV
{
    public class WeeklyReport : Evolution.EvolutionReport, IReport
    {
        private int rowIndex;
        private byte sheetIndex;

        private string tempMake;
        private string tempModel;
        private string tempVersion;


        public IEnumerable<Repository.CV.WEEKLY_CV> Data;
        
                

        public WeeklyReport()
        {
            if (!CheckExcel()) throw new Exception("Excel App not found");
            
            Evolution.EvolutionReportHeader evolutionHeader;

            SampleHeaders = new List<Evolution.EvolutionReportHeader>();

            evolutionHeader = new Evolution.EvolutionReportHeader("MSRP", "#,###", 0);
            evolutionHeader.ModelHeader.ModelSummaryFormula = "=IF(ISERROR(AVERAGE({0})),\" - \",AVERAGE({0}))";
            evolutionHeader.ModelHeader.ModelSummaryFormulaRanges = new int[,] { { 1, 0 }  };
            SampleHeaders.Add(evolutionHeader);

            evolutionHeader = new Evolution.EvolutionReportHeader("TP", "#,###", 1);
            evolutionHeader.ModelHeader.ModelSummaryFormula = "=IF(ISERROR(AVERAGE({0})),\" - \",AVERAGE({0}))";
            evolutionHeader.ModelHeader.ModelSummaryFormulaRanges = new int[,] { { 1, 0 } };
            SampleHeaders.Add(evolutionHeader);

            evolutionHeader = new Evolution.EvolutionReportHeader("Premium/Discount", "#.00%", 2);
            evolutionHeader.ModelHeader.ModelSummaryFormula = "=IF(ISERROR(AVERAGE({0})),\" - \",AVERAGE({0}))";
            evolutionHeader.ModelHeader.ModelSummaryFormulaRanges = new int[,] { { 1, 0 } };
            SampleHeaders.Add(evolutionHeader);
        }
        
        public WeeklyReport(IEnumerable<Repository.CV.WEEKLY_CV> data) : this()
        {
            if (data == null || data.Count() < 1)
            {
                throw new ArgumentException("Invalid WEEKLY_CV data");
            }
            else
            {
                Data = data;
            }
        }
        

        

        public override void GenerateEvolutionReport()
        {
            if (Data == null || Data.Count() < 1)
            {
                throw new InvalidOperationException("WEEKLY_CV Data is empty");
            }
            
            XlApp = new Excel.Application();
            XlBooks = XlApp.Workbooks;
            XlBook = XlBooks.Add(MisVal);
            XlSheets = XlApp.Sheets;
            XlSheets.Add(MisVal, MisVal, Data.Select(c => c.MAKE).Distinct().Count() - 1, MisVal);

            SampleDates = JATODbConverter.FromIntToDate(Data.Select(c => c.EXT_DATE).Distinct().OrderBy(c => c));
            VehDescRowMarkUpColumn = (int)constPostns.FirstSampleCol + SampleDates.Count() * SampleHeaders.Count;

            foreach (Repository.CV.WEEKLY_CV vehicle in Data)
            {
                if (vehicle.MAKE != tempMake)
                {
                    OnMakeChange(vehicle.MAKE);
                }

                if (vehicle.MODEL != tempModel)
                {
                    OnModelChange(vehicle.MODEL);
                }

                if (vehicle.VERSION != tempVersion)
                {
                    OnVersionChange(vehicle);
                }

                WriteVehicleData(vehicle, rowIndex);
            }

            FinalizeActiveWorksheet(rowIndex);
        }
        



        #region OnChange

        private void OnMakeChange(string newMakeName)
        {
            //If it is not the first make sheet
            if (tempMake != null)
            {
                FinalizeActiveWorksheet(rowIndex);
            }

            //Update temp make name and reset other temps
            tempMake = newMakeName;
            tempModel = null;
            tempVersion = null;
            SetupNewWorksheet(newMakeName, ++sheetIndex, VehDescRowMarkUpColumn + 1, (int)constPostns.FirstSampleRow + CalcAmountOfRowForVehiclesData(newMakeName));
            WriteWorksheetHeaders(newMakeName, "CV Weekly Report", "dd/MM/yyyy", "dd MMM yyyy");
            //Set row Idx to FirstSampleRow since Model will upate it 
                //regarding the possibility of coming from a version row
            rowIndex = (int)constPostns.FirstSampleRow - 1;
        }

        private void OnModelChange(string newModelName)
        {
            //If it is not a model of a new make sheet
            if (tempModel != null)
            {
                //Finish with the last version
                FillEmptyVehicleCells(rowIndex);
            }

            tempModel = newModelName;
            tempVersion = null;

            WriteModelHeader(newModelName, Data.Where(c => c.MAKE == tempMake && c.MODEL == tempModel).Select(c => new { c.MAKE, c.MODEL, c.VERSION }).Distinct().Count(), ++rowIndex);
        }

        private void OnVersionChange(Repository.CV.WEEKLY_CV vehicle)
        {
            //If it is not a version of a new model
            if (tempVersion != null)
            {
                FillEmptyVehicleCells(rowIndex);
            }

            //Update current version name
            tempVersion = vehicle.VERSION;
            //Jump to next row and mark-up this version row
            Matrix[++rowIndex, VehDescRowMarkUpColumn] = "x";
            //Write the vehicle's name
            Matrix[rowIndex, (int)constPostns.VehicleCol] = vehicle.VERSION == "OTHERS" ? "OTHERS" : tempVersion + " " + vehicle.PY.Trim().Substring(vehicle.PY.Length - 2) + "/" + vehicle.MY.Trim().Substring(vehicle.MY.Length - 2) + " " + vehicle.DOORS + "dr " + vehicle.BODY_TYPE;
        }

        #endregion


        private void WriteVehicleData(Repository.CV.WEEKLY_CV vehicle, int versionRow)
        {
            int curSampleWeekIndexOfCurVehicleVersion;
            int columnIndex;

            //Locate the version's TP sample week index inside of sampleWeekDates
            for (curSampleWeekIndexOfCurVehicleVersion = 0; curSampleWeekIndexOfCurVehicleVersion < SampleDates.Count(); curSampleWeekIndexOfCurVehicleVersion++)
            {
                if (JATODbConverter.FromIntToDate(vehicle.EXT_DATE) == SampleDates.ElementAt(curSampleWeekIndexOfCurVehicleVersion)) break;
            }

            columnIndex = (int)constPostns.FirstSampleCol + SampleHeaders.FirstOrDefault(h => h.HeaderName == "MSRP").Offset + curSampleWeekIndexOfCurVehicleVersion * SampleHeaders.Count;

            Matrix[versionRow, columnIndex++] = vehicle.MSRP_PLUS_OPC == null ? "-" : vehicle.MSRP_PLUS_OPC.ToString();
            Matrix[versionRow, columnIndex++] = vehicle.TP == null ? "-" : vehicle.TP.ToString();

            string msrpRange = XlUtils.XlCol(columnIndex - 2) + XlUtils.XlRow(versionRow);
            string tpRange = XlUtils.XlCol(columnIndex - 1) + XlUtils.XlRow(versionRow);

            Matrix[versionRow, columnIndex] = "=IF(ISERROR((" + tpRange + "-" + msrpRange + ")/" + msrpRange + "),\" - \",(" + tpRange + "-" + msrpRange + ")/" + msrpRange + ")";
        }
        
        private int CalcAmountOfRowForVehiclesData(string vehiclesMake)
        {
            var versionsGroup = (from car in Data
                                 where car.MAKE == vehiclesMake
                                 group car by new
                                 {
                                     MODEL = car.MODEL,
                                     VERSION = car.VERSION
                                 } into versiongroup
                                 select versiongroup);

            return versionsGroup.Count() + (from versn in versionsGroup
                                            group versn by new
                                            {
                                                MODEL = versn.Key.MODEL
                                            } into modelgroup
                                            select modelgroup).Count();
        }
        
    }
}
