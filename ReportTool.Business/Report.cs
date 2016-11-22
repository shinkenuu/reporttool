using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReportTool.Business
{
    public abstract class Report : IReport
    {
        #region Excel Variables

        protected readonly System.Reflection.Missing MisVal = System.Reflection.Missing.Value;

        protected Excel.Application XlApp;
        protected Excel.Workbooks XlBooks;
        protected Excel.Workbook XlBook;
        protected Excel.Sheets XlSheets;
        protected Excel.Worksheet XlWorksheet;

        #endregion
        
        protected string[,] Matrix;

        protected readonly string ReportsRootPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\ReportTool\Reports\";
        
        public abstract void GenerateReport();

        public void WriteToDisc()
        {
            if (!Directory.Exists(ReportsRootPath + this.ToString()))
            {
                Directory.CreateDirectory(ReportsRootPath + this.ToString());
            }

            string filePath = ReportsRootPath + this.ToString() + '\\' + this.ToString() + '_';

            for (int counter = 1; counter < int.MaxValue; counter++)
            {
                if (!File.Exists(filePath + counter.ToString() + ".xlsx"))
                {
                    filePath = filePath + counter.ToString() + ".xlsx";
                    break;
                }
            }

            try
            {
                XlBook.SaveAs(filePath, Excel.XlFileFormat.xlOpenXMLWorkbook, MisVal, MisVal, true, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlUserResolution, true, MisVal, MisVal, MisVal);
                XlBook.Close(true, MisVal, MisVal);
            }

            catch (Exception ex)
            {
                throw ex;
            }

            finally
            {
                CloseExcel();
            }
        }

        protected void CloseExcel()
        {
            try
            {
                if (XlWorksheet != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(XlWorksheet);
                    XlWorksheet = null;
                }

                if (XlSheets != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(XlSheets);
                    XlSheets = null;
                }

                if (XlBook != null)
                {
                    XlBook.Close(0);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(XlBook);
                    XlBook = null;
                }

                if (XlBooks != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(XlBooks);
                    XlBooks = null;
                }

                if (XlApp != null)
                {
                    XlApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(XlApp);
                    XlApp = null;
                }
            }

            catch (Exception)
            {
                XlWorksheet = null;
                XlSheets = null;
                XlBook = null;
                XlBooks = null;
                XlApp = null;
            }

            finally
            {
                GC.Collect();
            }
        }

        protected bool CheckExcel()
        {
            bool xlExists = false;

            try
            {
                XlApp = new Excel.Application();
                //If Excel isnt' installed in the host computer
                xlExists = XlApp != null;
            }
            catch (Exception)
            {
                XlApp = null;
                return false;
            }
            finally
            {
                CloseExcel();
            }

            return xlExists;
        }
        
        protected void WriteMatrixToExcel()
        {
            Excel.Range topLeftCellRange = (Excel.Range)XlWorksheet.Cells[1, 1];
            Excel.Range bottomRightCellRange = (Excel.Range)XlWorksheet.Cells[Matrix.GetLength(0), Matrix.GetLength(1)];
            Excel.Range range = XlWorksheet.get_Range(topLeftCellRange, bottomRightCellRange);

            try
            {
                range.Value = Matrix;
                range.Value = range.Value;
            }

            catch (Exception ex)
            {
                if (!ex.Message.StartsWith("Careful"))
                {
                    CloseExcel();
                    throw ex;
                }
            }
        }

    }
}
