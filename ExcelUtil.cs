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
using System.Data.OleDb;
using System.Diagnostics;
//using DocumentFormat.OpenXml.Packaging;

namespace CommonLib
{
    public static class ExcelUtil
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
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

        public static Excel.Workbook OpenWorkbook(string excelFileFullPath, out bool AlreadyOpened, bool WhenOpenCloseItFirst = false, bool readOnly = true)
        {
            Excel.Workbook wb = null;
            AlreadyOpened = false;
            bool anotherSameBaseFileNameOpened = false;
            //bool readOnlyFileOpenedButRequestEdit = false;
            Excel.Workbook readOnlyNoChange = null;
            Excel.Workbook readOnlyChanged = null;
            Excel.Workbook writeNoChange = null;
            Excel.Workbook writeChanged = null;

            string fileBaseName = Path.GetFileName(excelFileFullPath).ToUpper();

            foreach (Excel.Application item in GetAllExcelApplicationInstances())
            {
                log.Debug($"Excel.Application, item.Workbooks:  {item.Workbooks.Count}");
                foreach (Excel.Workbook eachWb in item.Workbooks)
                {
                    log.Debug($"eachWb: {eachWb.FullName}");
                    if (eachWb.FullName.ToUpper().Equals(excelFileFullPath.ToUpper()))
                    {
                        if (eachWb.ReadOnly)
                        {
                            if (!eachWb.Saved)
                            {
                                readOnlyChanged = eachWb;
                                break;
                            }
                            else
                            {
                                readOnlyNoChange = eachWb;
                                break;
                            }
                        }
                        else
                        {
                            if (!eachWb.Saved)
                            {
                                writeChanged = eachWb;
                                break;
                            }
                            else
                            {
                                writeNoChange = eachWb;
                                break;
                            }
                        }
                    }
                    else if (eachWb.Name.ToUpper().Equals(fileBaseName))
                    {
                        wb = null;
                        anotherSameBaseFileNameOpened = true;
                        break;
                    }
                }

                if (wb != null)
                    break;
            }

            if (readOnly)
            {
                if (readOnlyNoChange != null)
                {
                    AlreadyOpened = true;
                    return readOnlyNoChange;
                }
                else if (writeNoChange != null)
                {
                    AlreadyOpened = true;
                    return writeNoChange;
                }
                else if (readOnlyChanged != null)
                    return OpenExcelFileWithNewExcelApplication(excelFileFullPath, readOnly);
                else if (writeChanged != null)
                    return OpenExcelFileWithNewExcelApplication(excelFileFullPath, readOnly);
                else
                    return OpenExcelFileWithNewExcelApplication(excelFileFullPath, readOnly);
            }
            else
            { //for edit
                if (readOnlyNoChange != null)
                    return OpenExcelFileWithNewExcelApplication(excelFileFullPath, readOnly);
                else if (writeNoChange != null)
                {
                    AlreadyOpened = true;
                    return writeNoChange;
                }
                else if (writeChanged != null)
                {
                    DialogResult rs = MessageBox.Show($@"文件已经打开，并被修改，但没有保存，\r如果您想继续使用它，请点[是]，\r想重新打开它，丢失所做修改，请点[否]，\r停止处理，请点[取消]", "", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);
                    if (rs == DialogResult.Yes)
                    {
                        AlreadyOpened = true;
                        return writeChanged;
                    }
                    else if (rs == DialogResult.No)
                    {
                        writeChanged.Close(SaveChanges: false);
                        return OpenExcelFileWithNewExcelApplication(excelFileFullPath, readOnly);
                    }
                    else
                    {
                        throw new OperationCanceledException();
                    }
                }
                else if (readOnlyChanged != null)
                    return OpenExcelFileWithNewExcelApplication(excelFileFullPath, readOnly);
                else
                    return OpenExcelFileWithNewExcelApplication(excelFileFullPath, readOnly);
            }

            if (anotherSameBaseFileNameOpened && wb == null)
                return OpenExcelFileWithNewExcelApplication(excelFileFullPath, readOnly);

            return wb;
        }

        public static void SaveAndClose(this Excel.Workbook wb)
        {
            Excel.Application xlApp = wb.Parent;
            wb.CheckCompatibility = false;
            wb.Save();
            wb.CheckCompatibility = true;
            wb.Close();

            //GC.Collect();
            //GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(wb);
            //Marshal.FinalReleaseComObject(wb);

            if (xlApp.Workbooks.Count <= 0)
            {
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                //Marshal.FinalReleaseComObject(xlApp);
            }
        }
        public static void CloseWithoutSave(this Excel.Workbook wb)
        {

            Excel.Application xlApp = wb.Parent;
            wb.Close(SaveChanges: false);

            //GC.Collect();
            //GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(wb);
            //Marshal.FinalReleaseComObject(wb);

            if (xlApp.Workbooks.Count <= 0)
            {
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                //Marshal.FinalReleaseComObject(xlApp);
            }
        }

        //public static void CloseExcelApp(Excel.Application xlApp)
        //{
        //    xlApp.Visible = true;
        //    xlApp.Quit();
        //    //Marshal.ReleaseComObject(xlApp);
        //    Marshal.FinalReleaseComObject(xlApp);
        //}

        public static Excel.Workbook OpenExcelFileWithNewExcelApplication(string excelFileFullPath, bool readOnly = true)
        {
            Excel.Workbook wb = null;
            Excel.Application xlApp = GetDefaultFirstApplication();
            xlApp.Visible = true;
            wb = xlApp.Workbooks.Open(excelFileFullPath, UpdateLinks: false, ReadOnly: readOnly);

            if (!readOnly && wb.ReadOnly)
                MessageBox.Show("The excel file is opened as read only, but program is intended to write, there must be something wrong, please contact with IT support.", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return wb;
        }

        public static Excel.Workbook OpenWorkbook2(string excelFileFullPath, out bool alreadyOpened, bool WhenOpenCloseItFirst = false, bool readOnly = true)
        {
            Excel.Workbook wb = null;
            alreadyOpened = false;

            string fileBaseName = Path.GetFileName(excelFileFullPath).ToUpper();

            Excel.Application xlApp = GetDefaultFirstApplication();

            try
            {
                wb = xlApp.Workbooks.Open(excelFileFullPath, ReadOnly: readOnly, UpdateLinks: false);
            }
            catch (Exception ex)
            {
                throw;
            }

            return wb;
        }
        public static bool ExcelFileIsOpenedForEdit(string excelFileFullPath)
        {
            FileStream stream = null;
            try
            {
                stream = File.OpenWrite(excelFileFullPath);
            }
            catch (IOException ex)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }

        public static Excel.Workbook CreateWorkbook(string excelFileFullPath, string sheetName = null)
        {
            Excel.Workbook wb = null;

            if (File.Exists(excelFileFullPath))
            {
                DialogResult response = MessageBox.Show(text: "The file already exists, do you want to replace it with the new one?" + Environment.NewLine + excelFileFullPath
                    , caption: "File already exists", buttons: MessageBoxButtons.YesNoCancel, icon: MessageBoxIcon.Question);
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


        public static Excel.Workbook GetExactWorkbook(string excelFileFullPath)
        {
            Excel.Workbook wb = null;
            string fileBaseName = Path.GetFileName(excelFileFullPath).ToUpper();

            foreach (Excel.Application item in GetAllExcelApplicationInstances())
            {
                foreach (Excel.Workbook eachWb in item.Workbooks)
                {
                    if (eachWb.Name.ToUpper().Equals(fileBaseName))
                    {
                        wb = eachWb;
                        return wb;
                        break;
                    }
                }

                if (wb != null)
                    break;
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
                    fileFormat = Excel.XlFileFormat.xlOpenXMLWorkbook;
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
        //public static IEnumerable<Excel.Application> GetAllExcelApplicationInstances()
        //{
        //    Stopwatch stopwatch = Stopwatch.StartNew();
        //    log.Debug($"(Performance) GetAllExcelApplicationInstances started: {stopwatch.Elapsed}");

        //    HashSet<Excel.Application> result = new HashSet<Excel.Application>();

        //    IRunningObjectTable Rot;
        //    GetRunningObjectTable(0, out Rot);

        //    IEnumMoniker monikerEnumerator = null;
        //    Rot.EnumRunning(out monikerEnumerator);

        //    IntPtr iFetched = new IntPtr();
        //    IMoniker[] monikers = new IMoniker[1];

        //    IBindCtx ctx;
        //    CreateBindCtx(0, out ctx);

        //    while (monikerEnumerator.Next(1, monikers, iFetched) == 0)
        //    {
        //        string appName = "";
        //        dynamic wb = null;
        //        try
        //        {
        //            Guid iUnknown = new Guid("{00000000-0000-0000-C000-000000000046}");
        //            monikers[0].BindToObject(ctx, null, ref iUnknown, out wb);
        //            //appName = wb.Applicaiton.Name;
        //            appName = wb.Application.Name;
        //        }
        //        catch { }
        //        if (appName == "Microsoft Excel")
        //        {
        //            Excel.Application xlApp = wb.Application;

        //            if (!xlApp.Visible)
        //            {
        //                xlApp.Quit();
        //                Marshal.ReleaseComObject(xlApp);
        //            }
        //            else
        //            {
        //                if (!result.Contains(xlApp))
        //                {
        //                    stopwatch.Stop();
        //                    log.Debug($"(Performance) GetAllExcelApplicationInstances time consumed: {stopwatch.Elapsed}");
        //                    result.Add(xlApp);
        //                    //yield return xlApp;
        //                }
        //            }
        //        }
        //    }
        //    stopwatch.Stop();
        //    log.Debug($"(Performance) GetAllExcelApplicationInstances (total) time consumed: {stopwatch.Elapsed}");
        //    return result.ToList();
        //}
        public static IEnumerable<Excel.Application> GetAllExcelApplicationInstances()
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            log.Debug($"(Performance) GetAllExcelApplicationInstances started ...: {stopwatch.Elapsed}");

            HashSet<Excel.Application> result = new HashSet<Excel.Application>();

            IRunningObjectTable Rot;
            IEnumMoniker monikerEnumerator = null;
            IMoniker[] monikers = new IMoniker[1];

            IntPtr iFetched = new IntPtr();
            GetRunningObjectTable(0, out Rot);
            Rot.EnumRunning(out monikerEnumerator);

            IBindCtx ctx;
            CreateBindCtx(0, out ctx);
            Guid excelAppClsId = new Guid("{00000000-0000-0000-C000-000000000046}");

            while (monikerEnumerator.Next(1, monikers, iFetched) == 0)
            {
                string appName = "";
                dynamic wb = null;
                try
                {
                    monikers[0].BindToObject(ctx, null, ref excelAppClsId, out wb);
                    //appName = wb.Applicaiton.Name;
                    appName = wb.Application.Name;
                }
                catch { }
                log.Debug($"appName: {appName}");
                if (appName == "Microsoft Excel")
                {
                    Excel.Application xlApp = wb.Application;

                    log.Debug($"wb : {wb.Name}, readonly: {wb.ReadOnly}");

                    if (!xlApp.Visible)
                    {
                        log.Debug($"xlApp visible = false");
                        xlApp.Visible = true;
                        xlApp.Quit();
                        Marshal.ReleaseComObject(xlApp);
                    }
                    else
                    {
                        if (!result.Contains(xlApp))
                        {
                            //stopwatch.Stop();
                            //log.Debug($"(Performance) GetAllExcelApplicationInstances time consumed: {stopwatch.Elapsed}");
                            result.Add(xlApp);
                            yield return xlApp;
                        }
                    }
                }
            }
            stopwatch.Stop();
            log.Debug($"(Performance) GetAllExcelApplicationInstances (total) time consumed: {stopwatch.Elapsed}");
            //return result.ToList();
        }
        // Requires Using System.Runtime.InteropServices.ComTypes

        // Get all running instance by querying ROT

        public static List<object> GetRunningInstances(string[] progIds)
        {
            List<string> clsIds = new List<string>();

            // get the app clsid
            foreach (string progId in progIds)
            {
                Type type = Type.GetTypeFromProgID(progId);
                if (type != null)
                    clsIds.Add(type.GUID.ToString().ToUpper());
            }

            //Guid excelAppClsId = new Guid("{00000000-0000-0000-C000-000000000046}");
            //clsIds.Add(excelAppClsId.ToString().ToUpper());

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

                log.Debug($"displayName {displayName}");


                foreach (string clsId in clsIds)
                {
                    if (displayName.ToUpper().IndexOf(clsId) > 0)
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

        public static Excel.Application GetDefaultFirstApplication()
        {
            Excel.Application xlApp = null;

            try
            {
                xlApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch (Exception ex)
            {
                log.Error(ex.ToString());
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

            Excel.Application xlApp1 = null;
            int cnt = 0;

            foreach (Excel.Application xlApp2 in GetAllExcelApplicationInstances())
            {
                if (MultiInstanceThenGetTheDefault)
                {
                    xlApp1 = xlApp2;
                    xlApp1.Visible = true;
                    break;
                }
                else
                {
                    cnt++;
                    if (cnt > 1) break;
                }
            }
            if (!MultiInstanceThenGetTheDefault && cnt > 1)
                throw new Exception("Multiple Excel instances are running, please change your program to get the excel application directly " + Environment.NewLine + "Like moniker.bindtomonier(filename), or set this _xlapp");

            if (xlApp1 == null)
            {
                try
                {
                    xlApp1 = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch (Exception)
                {
                    xlApp1 = new Excel.Application();
                }
                finally
                {
                    xlApp1.Visible = true;
                }
            }

            return xlApp1;

            //Excel.Application xlApp = GetAllExcelApplicationInstances().FirstOrDefault();

            //if (xlApp ==null)
            //{
            //    xlApp = GetDefaultFirstApplication();
            //    xlApp.Visible = true;
            //}
            //else
            //{
            //    if (MultiInstanceThenGetTheDefault)
            //    {
            //        xlApp = GetDefaultFirstApplication();
            //        xlApp.Visible = true;
            //    }
            //    else
            //    {
            //        throw new Exception("Multiple Excel instances are running, please change your program to get the excel application directly " + Environment.NewLine + "Like moniker.bindtomonier(filename), or set this _xlapp");
            //    }
            //}
            //return xlApp;
        }
        #endregion

        #region Excel sheet names
        public static IEnumerable<string> GetExcelSheetNames(string excelFileFullPath)
        {
            //excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

            //bool bExcelFileIs2007Later = ExcelFileIs2007Later(excelFile);
            //bool bExcelFileIsOpen = ExcelFileIsOpen(excelFile);

            //if (bExcelFileIs2007Later && ! ExcelFileIsOpen)
            //{
            //    ExcelUtilOX.GetAllSheetNamesViaOpenXml
            //}
            //else
            //{

            //}


            //catch (System.IO.FileNotFoundException ex)
            //{
            //    log.Error(ex.ToString());
            //    throw ex;
            //}

            if (!File.Exists(excelFileFullPath))
            {
                //MessageBox.Show($@"文件不存在：{excelFileFullPath}");
                throw new System.IO.FileNotFoundException($@"文件不存在：{excelFileFullPath}");
            }


            IEnumerable<string> res = null;
            try
            {
                res = ExcelUtilOX.GetAllSheetNamesViaOpenXml(excelFileFullPath);
            }
            catch (FileFormatException ex)
            {
                log.Debug("The file is probably not Excel 2007 later. so program has to use tradition way.");
                log.Debug(ex.ToString());

                // use traditional way 



            }
            catch (IOException ex)
            {
                //file is already open

                log.Debug("The file is Excel 2007, but it is open, so program has to use tradition way.");
                log.Debug(ex.ToString());



            }
            catch (Exception ex)
            {

                throw;
            }
            return GetAllSheetNamesWithOleDb(excelFileFullPath);
        }


        public static string[] GetAllSheetNamesWithOleDb(string excelFile)
        {
            System.Data.DataTable dt = null;

            try
            {
                string connString = null;
                string extension = Path.GetExtension(excelFile);
                //switch (extension)
                //{
                //    case ".xls": //Excel 97-03.
                //        //connString = $@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={excelFile};Extended Properties=Excel 8.0;";
                //        connString = $@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={excelFile};Extended Properties='Excel 8.0;HDR=YES'";
                //        break;
                //    case ".xlsx": //Excel 07 or higher.
                //        connString = $@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = { excelFile }; Extended Properties ='Excel 8.0;HDR=YES'";
                //        break;
                //}

                //connString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={excelFile};Extended Properties=Excel 12.0;";
                connString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={excelFile};Extended Properties='Excel 8.0;HDR=YES'";
                connString = $@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={excelFile};Extended Properties='Excel 8.0;HDR=YES'";

                using (OleDbConnection objConn = new OleDbConnection(connString))
                {
                    objConn.Open();
                    dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    objConn.Close();
                }

                if (dt == null)
                    return null;

                String[] excelSheets = new String[dt.Rows.Count];
                int i = 0;
                foreach (System.Data.DataRow row in dt.Rows)
                {
                    excelSheets[i] = row["TABLE_NAME"].ToString();
                    i++;
                }

                return excelSheets;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        #endregion
    }
}
