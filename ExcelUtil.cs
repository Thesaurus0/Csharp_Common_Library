using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CommonLib
{
    public class ExcelUtil
    {
        private static Excel.Application _ExcelApp = GetExcelApplication();

        #region GetSheetBySheetName
        public static Excel.Worksheet GetSheetBySheetName(string shtName, Excel.Workbook wb = null)
        {
            Excel.Worksheet shtOut = null;

            wb = SetParamWorkbookToAcitveWorkbookWhenItIsNull(wb);

            foreach (Excel.Worksheet item in wb.Worksheets)
            {
                if (shtName.ToUpper().Equals(item.Name.ToUpper()))
                {
                    shtOut = item;
                    break;
                }
            }
            return shtOut;
        }
        #endregion

        #region GetSheetByCodeName
        public static Excel.Worksheet GetSheetByCodeName(string shtCodeName, Excel.Workbook wb = null)
        {
            Excel.Worksheet shtOut = null;

            wb = SetParamWorkbookToAcitveWorkbookWhenItIsNull(wb);

            foreach (Excel.Worksheet item in wb.Worksheets)
            {
                if (shtCodeName.ToUpper().Equals(item.Name.ToUpper()))
                {
                    shtOut = item;
                    break;
                }
            }
            return shtOut;
        }
        #endregion
        #region GetRangeByStartEndPos
        public static Excel.Range GetRangeByStartEndPos(Excel.Worksheet sht, long StartRow, long StartCol, long EndRow, long EndCol)
        {
            Excel.Range rg = null;

            if (StartRow > EndRow)
                throw new ArgumentException("StartRow > EndRow");
            if (StartCol > EndCol)
                throw new ArgumentException("StartCol > EndCol");

            rg = sht.Range[sht.Cells[StartRow, StartCol], sht.Cells[EndRow, EndCol]];
            return rg;
        }
        #endregion

        #region SetParamWbToAcitveWorkbookWhenItIsNull
        private static Excel.Workbook SetParamWorkbookToAcitveWorkbookWhenItIsNull(Excel.Workbook wb)
        {
            if (wb != null)
            {
                return wb;
            }
            else
            {
                if (_ExcelApp.Workbooks.Count > 0)
                {
                    return _ExcelApp.ActiveWorkbook;
                }
                else
                {
                    throw new ArgumentException("Parameter wb is null, and there is no actuve workbook available (in SetParamWbToAcitveWorkbookWhenItIsNull)");
                }
            }
        }
        #endregion

        #region GetExcelApplication

        private static Excel.Application GetExcelApplication()
        {
            Excel.Application xlApp = null;

            try
            {
                xlApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            }
            catch (Exception excp)
            {
                xlApp = new Excel.Application();
                //throw excp;
            }
            finally
            {
                xlApp.Visible = true;
            }

            xlApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

            return xlApp;
        }
        #endregion


    }
}
