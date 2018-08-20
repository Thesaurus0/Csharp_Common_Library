using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CommonLib
{
    public class ExcelUtil
    {
        private const string SHEET_CODE_NAME = "SHEET_CODE_NAME";
        private static Excel.Application _ExcelApp = GetExcelApplication();

        public static Excel.Application ExcelAPP
        {
            get { return _ExcelApp; }
        }

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
            wb = SetParamWorkbookToAcitveWorkbookWhenItIsNull(wb);

            foreach (Excel.Worksheet item in wb.Worksheets)
            {
                if (shtCodeName.ToUpper().Equals(GetSheetCodeName(item).ToUpper()))
                {
                    return item;
                }
            }
            return null;
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
        private static Excel.Application GetExcelApplication(bool MultiInstanceThenGetTheDefault = false)
        {
            Excel.Application xlApp = null;

            if (!MultiInstanceThenGetTheDefault)
            {
                try
                {
                    MultipleExcelApplicationIsRunning();
                }
                catch (Exception excp)
                {
                    System.Windows.Forms.MessageBox.Show(excp.Message);
                    throw excp;
                }
            }

            try
            {
                xlApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
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

            return xlApp;
        }
        #endregion

        #region GetSheetCodeName
        public static string GetSheetCodeName(Excel.Worksheet sht)
        {
            return GetSheetProperty(sht, SHEET_CODE_NAME);
        }
        #endregion

        #region SetSheetCodeName
        public static void SetSheetCodeName(Excel.Worksheet sht, string shtCodeName)
        {
            SetSheetProperty(sht, SHEET_CODE_NAME, shtCodeName);
        }
        #endregion

        #region SetSheetProperty
        private static void SetSheetProperty(Excel.Worksheet sht, string propertyName, string propertyValue)
        {
            bool existing = false;

            foreach (Excel.CustomProperty item in sht.CustomProperties)
            {
                if (propertyName.ToUpper().Equals(((string)item.Name).ToUpper()))
                {
                    existing = true;
                    item.Value = propertyValue;
                    break;
                }
            }

            if (!existing)
                sht.CustomProperties.Add(propertyName, propertyValue);
        }
        #endregion

        #region GetSheetProperty
        private static string GetSheetProperty(Excel.Worksheet sht, string propertyName)
        {
            foreach (Excel.CustomProperty item in sht.CustomProperties)
            {
                if (propertyName.ToUpper().Equals(((string)item.Name).ToUpper()))
                {
                    return (string)item.Value;
                }
            }
            return null;
        }
        #endregion

        //private static Excel.Application GetExcelApplications()
        //{
        //    Excel.Application xlApp = null;

        //    [DllImport("ole32.dll")]
        //static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        //[DllImport("ole32.dll")]

        //public static extern void GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

        //    return xlApp;
        //}

        #region MultipleExcelApplicationIsRunning
        public static bool MultipleExcelApplicationIsRunning(bool ThrowsExceptionAtOnce = true)
        {
            int instanceCnt = 0;

            List<object> instances = GetRunningInstances(new string[1] { "Excel.Application" });
            instanceCnt = instances.Count;
            if (instanceCnt <= 1)
                return false;

            Excel.Application xlApp = null;
            foreach (var item in instances)
            {
                bool toRelease = false;

                if (item is Excel.Application)
                { 
                    xlApp = (item as Excel.Application);
                    xlApp.Visible = true;
                }

                if (xlApp.Visible == true)
                {
                    if (xlApp.ActiveWorkbook.Name == null)
                        toRelease = true;
                }
                else
                {
                    toRelease = true;
                }

                if (toRelease)
                {
                    instanceCnt--;
                    if (instanceCnt > 0)
                    {
                        xlApp.Quit();
                        Marshal.ReleaseComObject(xlApp);
                    }
                    else
                    {
                        return false;
                    }
                }
            }


            if (instanceCnt > 1)
            {
                if (ThrowsExceptionAtOnce)
                    throw new Exception("Multiple Excel instance is running, you may now try to get the activeworkook, so that program cannot determine which instance's activeworkook you want. ");
            }

            return instanceCnt > 1;
        }
        #endregion

        [DllImport("ole32.dll")]
        static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        [DllImport("ole32.dll")]
        public static extern void GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

        #region GetRunningInstances

        // Requires Using System.Runtime.InteropServices.ComTypes
        // Get all running instance by querying ROT
        private static List<object> GetRunningInstances(string[] progIds)
        {
            List<string> clsIds = new List<string>();
            // get the app clsid
            foreach (string progId in progIds)
            {
                Type type = Type.GetTypeFromProgID(progId);

                if (type != null)
                    clsIds.Add(type.GUID.ToString().ToUpper());
            }
            // get Running Object Table ...
            IRunningObjectTable Rot = null;
            GetRunningObjectTable(0, out Rot);
            if (Rot == null)
                return null;
            // get enumerator for ROT entries
            IEnumMoniker monikerEnumerator = null;

            Rot.EnumRunning(out monikerEnumerator);
            if (monikerEnumerator == null)
                return null;
            monikerEnumerator.Reset();

            List<object> instances = new List<object>();

            IntPtr pNumFetched = new IntPtr();

            IMoniker[] monikers = new IMoniker[1];

            // go through all entries and identifies app instances
            while (monikerEnumerator.Next(1, monikers, pNumFetched) == 0)
            {
                IBindCtx bindCtx;
                CreateBindCtx(0, out bindCtx);
                if (bindCtx == null)
                    continue;

                string displayName;
                monikers[0].GetDisplayName(bindCtx, null, out displayName);

                foreach (string clsId in clsIds)
                {
                    if (displayName.ToUpper().IndexOf(clsId.ToUpper()) > 0)
                    {
                        object ComObject;

                        Rot.GetObject(monikers[0], out ComObject);

                        if (ComObject == null)
                            continue;
                        instances.Add(ComObject);
                        break;
                    }
                }
            }
            return instances;
        }
        #endregion
    }
}
