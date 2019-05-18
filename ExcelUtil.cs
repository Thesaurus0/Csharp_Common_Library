using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Windows.Forms;

namespace CommonLib
{
    public static class ExcelUtil
    {
        private const string SHEET_CODE_NAME = "SHEET_CODE_NAME";

        #region windows API
        [DllImport("ole32.dll")]
        public static extern void GetRunningObjectTable(int reserved, out IRunningObjectTable prot);
        [DllImport("ole32.dll")]
        static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);
        #endregion

        #region _ExcelApp
        private static Excel.Application _ExcelApp = null;
        public static Excel.Application ExcelAPP
        {
            get { return _ExcelApp; }
            set { _ExcelApp = value; }
        }
        #endregion

        #region workbook operation

        public static Excel.Workbook OpenWorkbook(string excelFileFullPath, out bool AlreadyOpened, bool WhenOpenCloseItFirst = false, bool ReadOnly = true )
        {
            Excel.Workbook wb = null;
            AlreadyOpened = false;

            wb = GetWorkbook(excelFileFullPath);
            if (wb != null)
            {
                if (wb.FullName.ToUpper().Equals(excelFileFullPath.Trim().ToUpper()))
                    return wb;
                else
                {

                }

            }
            else
            {

            }

            return wb;
        }

        public static Excel.Workbook CreateWorkbook(string excelFileFullPath, string sheetName = null)
        {
            Excel.Workbook wb = null;

            if (File.Exists(excelFileFullPath))
            {
                DialogResult response = MessageBox.Show(text:"The file already exists, do you want to replace it with the new one?" + Environment.NewLine + excelFileFullPath 
                    , caption:"File already exists", buttons: MessageBoxButtons.YesNoCancel, icon: MessageBoxIcon.Question);
                if (response != DialogResult.Yes)
                    throw new Exception("User abort");
            }

            _ExcelApp.SheetsInNewWorkbook = 1;
            wb = _ExcelApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            _ExcelApp.ActiveWindow.DisplayGridlines = false;

            wb.Worksheets[1].Name = sheetName ?? Path.GetFileNameWithoutExtension(excelFileFullPath);

            Excel.XlFileFormat fileFormat = GetExcelFileFormatByFileExtension(excelFileFullPath);
            wb.SaveAs(Filename: excelFileFullPath, FileFormat: fileFormat);
            return wb;
        }
        public static Excel.Workbook GetActiveWorkbook()
        {
            return GetExcelApplication()?.ActiveWorkbook;
        }


        public static Excel.Workbook GetWorkbook(string excelFileFullPath)
        {
            Excel.Workbook wb = null;

            foreach (Excel.Application item in GetAllExcelApplicationInstances())
            {
                if (WorkbookExistsInExcelApplication(excelFileFullPath, out wb, item))
                {
                    break;
                }
            }
            return null;
        }


        private static bool WorkbookExistsInExcelApplication(string excelFileName, out Excel.Workbook wb, Excel.Application xlApp = null)
        {
            string fileBaseName = Path.GetFileName(excelFileName).ToUpper();
            wb = null;

            if (xlApp == null)
                xlApp = _ExcelApp ?? GetDefaultFirstApplication();

            foreach (Excel.Workbook eachWb in xlApp.Workbooks)
            {
                if (eachWb.Name.ToUpper().Equals(fileBaseName))
                {
                    wb = eachWb;
                    return true;
                }
            }

            return false;
        }
        private static bool WorkbookExistsInExcelApplication(string excelFileName, Excel.Application xlApp = null)
        {
           Excel.Workbook wb = null;
            return WorkbookExistsInExcelApplication(excelFileName, out wb, xlApp);
        }


        private static Excel.XlFileFormat GetExcelFileFormatByFileExtension(string excelFileFullPath)
        {
            Excel.XlFileFormat fileFormat = default(Excel.XlFileFormat);
            string ext = Path.GetExtension(excelFileFullPath).ToUpper();

            switch (ext)
            {
                case ".CSV":
                    fileFormat = Excel.XlFileFormat.xlCSV;
                    break;
                case ".XLS":
                    fileFormat = Excel.XlFileFormat.xlExcel8;
                    break;
                case ".XLSX":
                    fileFormat = Excel.XlFileFormat.xlOpenXMLWorkbook ;
                    break;
                case ".XLSM":
                    fileFormat = Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled;
                    break;
                case ".XLSB":
                    fileFormat = Excel.XlFileFormat.xlExcel12;
                    break;
                case ".TXT":
                    fileFormat = Excel.XlFileFormat.xlCurrentPlatformText;
                    break;
                case ".PRN":
                    fileFormat = Excel.XlFileFormat.xlTextPrinter;
                    break;
                default:
                    throw new Exception("File extension is invalid :" + ext.ToLower());
                    //break;
            }
            return fileFormat;
        }
        #endregion

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
        public static Excel.Range GetRangeByStartEndPos(this Excel.Worksheet sht, long StartRow, long StartCol, long EndRow, long EndCol)
        {
            Excel.Range rg = null;

            if (StartRow > EndRow)
                throw new ArgumentException("StartRow > EndRow");
            if (StartCol > EndCol)
                throw new ArgumentException("StartCol > EndCol");

            rg = sht.Range[sht.Cells[StartRow, StartCol], sht.Cells[EndRow, EndCol]];
            return rg;
        }

        private static Excel.Workbook SetParamWorkbookToAcitveWorkbookWhenItIsNull(Excel.Workbook wb)
        {
            if (wb != null)
            {
                return wb;
            }
            else
            {
                if (_ExcelApp.Workbooks.Count > 0)
                    return _ExcelApp.ActiveWorkbook;
                else
                    throw new ArgumentException("Parameter wb is null, and there is no actuve workbook available (in SetParamWbToAcitveWorkbookWhenItIsNull)");
            }
        }


        public static string GetSheetCodeName(Excel.Worksheet sht)
        {
            return GetSheetProperty(sht, SHEET_CODE_NAME);
        }

        public static void SetSheetCodeName(Excel.Worksheet sht, string shtCodeName)
        {
            SetSheetProperty(sht, SHEET_CODE_NAME, shtCodeName);
        }

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

        #region comment out
        //public static bool MultipleExcelApplicationIsRunning(bool ThrowsExceptionAtOnce = true)
        //{
        //    int instanceCnt = 0;

        //    List<object> instances = GetRunningInstances(new string[1] { "Excel.Application" });
        //    instanceCnt = instances.Count;
        //    if (instanceCnt <= 1)
        //        return false;

        //    Excel.Application xlApp = null;
        //    foreach (var item in instances)
        //    {
        //        bool toRelease = false;

        //        if (item is Excel.Application)
        //        { 
        //            xlApp = (item as Excel.Application);
        //            xlApp.Visible = true;
        //        }

        //        if (xlApp.Visible == true)
        //        {
        //            if (xlApp.ActiveWorkbook.Name == null)
        //                toRelease = true;
        //        }
        //        else
        //        {
        //            toRelease = true;
        //        }

        //        if (toRelease)
        //        {
        //            instanceCnt--;
        //            if (instanceCnt > 0)
        //            {
        //                xlApp.Quit();
        //                Marshal.ReleaseComObject(xlApp);
        //            }
        //            else
        //            {
        //                return false;
        //            }
        //        }
        //    }


        //    if (instanceCnt > 1)
        //    {
        //        if (ThrowsExceptionAtOnce)
        //            throw new Exception("Multiple Excel instance is running, you may now try to get the activeworkook, so that program cannot determine which instance's activeworkook you want. ");
        //    }

        //    return instanceCnt > 1;
        //}



        // Requires Using System.Runtime.InteropServices.ComTypes
        // Get all running instance by querying ROT
        //private static List<object> GetRunningInstances(string[] progIds)
        //{
        //    List<string> clsIds = new List<string>();
        //    // get the app clsid
        //    foreach (string progId in progIds)
        //    {
        //        Type type = Type.GetTypeFromProgID(progId);

        //        if (type != null)
        //            clsIds.Add(type.GUID.ToString().ToUpper());
        //    }
        //    // get Running Object Table ...
        //    IRunningObjectTable Rot = null;
        //    GetRunningObjectTable(0, out Rot);
        //    if (Rot == null)
        //        return null;
        //    // get enumerator for ROT entries
        //    IEnumMoniker monikerEnumerator = null;

        //    Rot.EnumRunning(out monikerEnumerator);
        //    if (monikerEnumerator == null)
        //        return null;
        //    monikerEnumerator.Reset();

        //    List<object> instances = new List<object>();

        //    IntPtr pNumFetched = new IntPtr();

        //    IMoniker[] monikers = new IMoniker[1];

        //    // go through all entries and identifies app instances
        //    while (monikerEnumerator.Next(1, monikers, pNumFetched) == 0)
        //    {
        //        IBindCtx bindCtx;
        //        CreateBindCtx(0, out bindCtx);
        //        if (bindCtx == null)
        //            continue;

        //        string displayName;
        //        monikers[0].GetDisplayName(bindCtx, null, out displayName);

        //        foreach (string clsId in clsIds)
        //        {
        //            if (displayName.ToUpper().IndexOf(clsId.ToUpper()) > 0)
        //            {
        //                object ComObject;

        //                Rot.GetObject(monikers[0], out ComObject);

        //                if (ComObject == null)
        //                    continue;
        //                instances.Add(ComObject);

        //                //var v = typeof(   );
        //                break;
        //            }
        //        }
        //    }
        //    return instances;
        //}
        #endregion

        #region Excel Application Operation
        private static List<Excel.Application> GetAllExcelApplicationInstances()
        {
            HashSet<Excel.Application> result = new HashSet<Excel.Application>();

            IRunningObjectTable Rot;
            GetRunningObjectTable(0, out Rot);

            IEnumMoniker monikerEnumerator = null;
            Rot.EnumRunning(out monikerEnumerator);

            IntPtr iFetched = new IntPtr();
            IMoniker[] monikers = new IMoniker[1];


            IBindCtx ctx;
            CreateBindCtx(0, out ctx);

            while (monikerEnumerator.Next(1, monikers, iFetched) == 0)
            {
                string appName = "";
                dynamic wb = null;
                try
                {
                    Guid iUnknown = new Guid("{00000000-0000-0000-C000-000000000046}");
                    monikers[0].BindToObject(ctx, null, ref iUnknown, out wb);
                    //appName = wb.Applicaiton.Name;
                    appName = wb.Application.Name;
                }
                catch { }
                if (appName == "Microsoft Excel")
                {
                    Excel.Application xlApp = wb.Application;

                    if (!xlApp.Visible)
                    {
                        xlApp.Quit();
                        Marshal.ReleaseComObject(xlApp);
                    }
                    else
                    {
                        if (!result.Contains(xlApp))
                        {
                            result.Add(xlApp);
                        }
                    }
                }
            }
            return result.ToList();
        }

        private static Excel.Application GetDefaultFirstApplication()
        {
            Excel.Application xlApp = null;

            try
            {
                xlApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch (Exception)
            {
                xlApp = new Excel.Application();
            }
            finally
            {
                xlApp.Visible = true;
            }
            return xlApp;
        }
        private static Excel.Application GetExcelApplication(bool MultiInstanceThenGetTheDefault = false)
        {
            if (_ExcelApp != null)
                return _ExcelApp;

            int excelAppInstanceNum = GetAllExcelApplicationInstances().Count;

            Excel.Application xlApp = null;

            if (excelAppInstanceNum <= 1)
            {
                xlApp = GetDefaultFirstApplication();
                xlApp.Visible = true;
            }
            else
            {
                if (MultiInstanceThenGetTheDefault)
                {
                    xlApp = GetDefaultFirstApplication();
                    xlApp.Visible = true;
                }
                else
                {
                    throw new Exception("Multiple Excel instances are running, please change your program to get the excel application directly " + Environment.NewLine + "Like moniker.bindtomonier(filename), or set this _xlapp");
                }
            }
            return xlApp;
        }
        #endregion
    }
}
