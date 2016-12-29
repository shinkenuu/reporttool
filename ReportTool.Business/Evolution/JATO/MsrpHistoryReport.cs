﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using XlUtils = JATO.Common.XlInterop.Utils;
using JatoDbConverter = JATO.Common.Database.Converter;

namespace ReportTool.Business.Evolution.JATO
{
    public class MsrpHistoryReport : EvolutionReport, IReport
    {
        private int rowIndex;
        private byte sheetIndex;

        private string tempMake;
        private string tempModel;
        private string tempVersion;

        public IEnumerable<Repository.RT.MSRP_HISTORY> Data;

        public MsrpHistoryReport()
        {
            if (!CheckExcel()) throw new Exception("Excel App not found");

            Evolution.EvolutionReportHeader evolutionHeader;

            SampleHeaders = new List<Evolution.EvolutionReportHeader>();

            evolutionHeader = new Evolution.EvolutionReportHeader("MSRP", "#,###", 0);
            evolutionHeader.ModelHeader.ModelSummaryFormula = "=IF(ISERROR(AVERAGE({0})), \" - \",AVERAGE({0}))";
            evolutionHeader.ModelHeader.ModelSummaryFormulaRanges = new int[,] { { 1, 0 } };
            SampleHeaders.Add(evolutionHeader);

            evolutionHeader = new Evolution.EvolutionReportHeader("Premium/Discount", "#.00%", 1);
            evolutionHeader.ModelHeader.ModelSummaryFormula = "=IF(ISERROR(SUMPRODUCT({0},{1})/SUM({1})), \" - \",SUMPRODUCT({0},{1})/SUM({1}))";
            evolutionHeader.ModelHeader.ModelSummaryFormulaRanges = new int[,] { { 1, 0 }, { 1, -1 } };
            SampleHeaders.Add(evolutionHeader);

            ////MSRP weighted with Volumes = "=IF(ISERROR(SUMPRODUCT(" + msrpColumnRange + "," + volumesRange + ")/SUM(" + volumesRange + ")), \" - \",SUMPRODUCT(" + msrpColumnRange + "," + volumesRange + ")/SUM(" + volumesRange + "))";
            ////TP weighted with Volumes = "=IF(ISERROR(SUMPRODUCT(" + tpColumnRange + "," + volumesRange + ")/SUM(" + volumesRange + ")), \" - \",SUMPRODUCT(" + tpColumnRange + "," + volumesRange + ")/SUM(" + volumesRange + "))";
            ////Sum of Volumes = "=IF(ISERROR(SUM(" + volumesRange + ")), \" - \", SUM(" + volumesRange + "))";
            ////Premium/Discount weighted with Volumes = "=IF(ISERROR(SUMPRODUCT(" + discountRange + "," + volumesRange + ")/SUM(" + volumesRange + ")), \" - \",SUMPRODUCT(" + discountRange + "," + volumesRange + ")/SUM(" + volumesRange + "))";

            //Total Discount / Premium % = "=IF(ISERROR(SUM(" + columnLetter + XlRow(modelHeaderRow + 1).ToString() + ":" + columnLetter + XlRow(modelHeaderRow + amountOfVehiclesOfCurModel).ToString() + ")), \" - \",SUM(" + columnLetter + XlRow(modelHeaderRow + 1).ToString() + ":" + columnLetter + XlRow(modelHeaderRow + amountOfVehiclesOfCurModel).ToString() + "))";
            //Total Volume = "=IF(ISERROR(SUM(" + columnLetter + XlRow(modelHeaderRow + 1).ToString() + ":" + columnLetter + XlRow(modelHeaderRow + amountOfVehiclesOfCurModel).ToString() + ")), \" - \",SUM(" + columnLetter + XlRow(modelHeaderRow + 1).ToString() + ":" + columnLetter + XlRow(modelHeaderRow + amountOfVehiclesOfCurModel).ToString() + "))";
        }

        public MsrpHistoryReport(IEnumerable<Repository.RT.MSRP_HISTORY> data) : this()
        {
            if (data == null || data.Count() < 1)
            {
                throw new ArgumentException("Invalid MONTHLY_CV data");
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
                throw new InvalidOperationException("MSRP_HISTORY Data is empty");
            }

            XlApp = new Excel.Application();
            XlBooks = XlApp.Workbooks;
            XlBook = XlBooks.Add(MisVal);
            XlSheets = XlApp.Sheets;
            XlSheets.Add(MisVal, MisVal, Data.Select(c => c.MAKE).Distinct().Count() - 1, MisVal);

            SampleDates = JatoDbConverter.FromIntToDate(Data.Select(c => c.DATADATE).Distinct().OrderBy(c => c));
            VehDescRowMarkUpColumn = (int)constPostns.FirstSampleCol + SampleDates.Count() * SampleHeaders.Count;

            foreach (Repository.RT.MSRP_HISTORY vehicle in Data)
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
            WriteWorksheetHeaders(newMakeName, "MSRP History Report", "dd/MM/yyyy", "dd/MM/yyyy");
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

        private void OnVersionChange(Repository.RT.MSRP_HISTORY vehicle)
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
            Matrix[rowIndex, (int)constPostns.VehicleCol] = FormatVehicleName(vehicle);
        }

        #endregion







        private void WriteVehicleData(Repository.RT.MSRP_HISTORY vehicle, int versionRow)
        {
            int sampleDateIdx = SampleDates.ToList().IndexOf(JatoDbConverter.FromIntToDate(vehicle.DATADATE));
            int columnIndex = (int)constPostns.FirstSampleCol + sampleDateIdx * SampleHeaders.Count;
            
            if(vehicle.MSRP == null)
            {
                Matrix[versionRow, columnIndex++] = "\" - \"";
                Matrix[versionRow, columnIndex] = "\" - \"";
            }

            else
            {
                //MSRP
                Matrix[versionRow, columnIndex++] = vehicle.MSRP.ToString();
                string msrpCell = XlUtils.XlCol(columnIndex) + XlUtils.XlRow(versionRow).ToString();

                //Premium/Discount
                string lastMsrpChangeCell = LastMsrpChangeCell(versionRow, columnIndex - (SampleHeaders.FirstOrDefault(s => s.HeaderName == "MSRP").Offset + 1));
                Matrix[versionRow, columnIndex++] = "=IF(ISERROR((" + msrpCell + "/" + lastMsrpChangeCell + ")-1), \" - \",(" + msrpCell + "/" + lastMsrpChangeCell + ")-1)";
            }
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

        private string LastMsrpChangeCell(int versionRow, int currentMsrpColumn)
        {
            int headersCount = SampleHeaders.Max(s => s.Offset) + 1;

            for(int tgtColumn = currentMsrpColumn; tgtColumn >= (int)constPostns.FirstSampleCol; tgtColumn -= headersCount)
            {
                if (Matrix[versionRow, tgtColumn] != "?" || Matrix[versionRow, tgtColumn] != "\" - \"")
                {
                    return XlUtils.XlCol(tgtColumn) + XlUtils.XlRow(versionRow);
                } 
            }
            
            return null;
        }

        private string FormatVehicleName(Repository.RT.MSRP_HISTORY vehicle)
        {
            if (vehicle.VERSION == "OTHERS") return "OTHERS";

            string vehicleName = "[" + vehicle.UID.ToString() + "] " + vehicle.VERSION;

            if (vehicle.PY == null) vehicleName += " ??/";
            else vehicleName += " " + vehicle.PY.Trim().Substring(Math.Max(0, vehicle.PY.Trim().Length - 2)) + "/";

            if (vehicle.MY == null) vehicleName += "??";
            else vehicleName += vehicle.MY.Trim().Substring(Math.Max(0, vehicle.MY.Trim().Length - 2));

            if (vehicle.DOORS == null) vehicleName += " ";
            else vehicleName += " " + vehicle.DOORS + "dr";

            if (vehicle.BODY_TYPE == null) return vehicleName;
            else return vehicleName + " " + vehicle.BODY_TYPE;
        }
        
    }
}
