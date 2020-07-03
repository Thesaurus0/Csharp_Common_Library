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
using System.Data;
using System.Reflection;
//using DocumentFormat.OpenXml.Packaging;
using ADODB;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CommonLib
{
    //[Flags]
    //public enum ExcelFileOpenStatus
    //{
    //    None = 0,
    //    //FileNotOpen = None,
    //    ReadOnlyAndNoChange = 2,
    //    ReadOnlyButChangeNotSave = 4,
    //    WriteAndNoChange = 8,
    //    WriteButChangeNotSave = 16,
    //    AnotherSameFileNameOpened = 32,
    //    //FileIsOpen = ReadOnlyAndNoChange | ReadOnlyButChangeNotSave  | WriteAndNoChange | WriteButChangeNotSave  
    //}

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
        //private static Excel.Application _ExcelApp = null;
        //public static Excel.Application ExcelAPP
        //{
        //    get { return _ExcelApp; }
        //    set { _ExcelApp = value; }
        //}
        #endregion

        #region workbook operation
        public static Excel.Workbook OpenWorkbook(string excelFileFullPath, out bool AlreadyOpened, bool readOnly = true)
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
                //log.Debug($"Excel.Application, item.Workbooks:  {item.Workbooks.Count}");
                foreach (Excel.Workbook eachWb in item.Workbooks)
                {
                    //log.Debug($"eachWb: {eachWb.FullName}");
                    if (eachWb.FullName.ToUpper().Equals(excelFileFullPath.ToUpper()))
                    {
                        if (eachWb.ReadOnly)
                        {
                            if (!eachWb.Saved)
                            {
                                log.Debug($"workbook was opened readonly, with changes made.{eachWb.FullName}");
                                readOnlyChanged = eachWb;
                                break;
                            }
                            else
                            {
                                log.Debug($"workbook was opened readonly, without any change.{eachWb.FullName}");
                                readOnlyNoChange = eachWb;
                                break;
                            }
                        }
                        else
                        {
                            if (!eachWb.Saved)
                            {
                                log.Debug($"workbook was opened in edit mode, with changes made.{eachWb.FullName}");
                                writeChanged = eachWb;
                                break;
                            }
                            else
                            {
                                log.Debug($"workbook was opened in edit mode, without any change.{eachWb.FullName}");
                                writeNoChange = eachWb;
                                break;
                            }
                        }
                    }
                    else if (eachWb.Name.ToUpper().Equals(fileBaseName))
                    {
                        log.Debug($"same workbook name was found : {eachWb.FullName}");
                        wb = null;
                        anotherSameBaseFileNameOpened = true;
                        break;
                    }
                }

                if (wb != null)
                    break;
            }

            if (anotherSameBaseFileNameOpened && wb == null)
            {
                //return OpenExcelFileWithNewExcelApplication(excelFileFullPath, readOnly);
                return OpenExcelFileWithDefaultExcelApplication(excelFileFullPath, readOnly);
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
                    return OpenExcelFileWithDefaultExcelApplication(excelFileFullPath, readOnly);
            }
            else
            { //for edit
                if (writeNoChange != null)
                {
                    AlreadyOpened = true;
                    return writeNoChange;
                }
                else if (writeChanged != null)
                {
                    DialogResult rs = MessageBox.Show($@"文件已经打开，并被修改，但没有保存，{Environment.NewLine}如果您想继续使用它，请点[是]，{Environment.NewLine}想重新打开它，丢失所做修改，请点[否]，{Environment.NewLine}停止处理，请点[取消]", "", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);
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
                else if (readOnlyNoChange != null)
                    return OpenExcelFileWithNewExcelApplication(excelFileFullPath, readOnly);
                else if (readOnlyChanged != null)
                    return OpenExcelFileWithNewExcelApplication(excelFileFullPath, readOnly);
                else
                    return OpenExcelFileWithDefaultExcelApplication(excelFileFullPath, readOnly);
            }

            return wb;
        }
        //public static Excel.Workbook OpenWorkbook2(string excelFileFullPath, out bool alreadyOpened, bool WhenOpenCloseItFirst = false, bool readOnly = true)
        //{
        //    Excel.Workbook wb = null;
        //    alreadyOpened = false;

        //    //string fileBaseName = Path.GetFileName(excelFileFullPath).ToUpper();

        //    //Excel.Application xlApp = GetDefaultFirstApplication();

        //    //try
        //    //{
        //    //    wb = xlApp.Workbooks.Open(excelFileFullPath, ReadOnly: readOnly, UpdateLinks: false);
        //    //}
        //    //catch (Exception ex)
        //    //{
        //    //    throw;
        //    //}

        //    return wb;
        //}
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

        public static void SaveButNotClose(this Excel.Workbook wb)
        {
            wb.CheckCompatibility = false;
            wb.Save();
            wb.CheckCompatibility = true;
            wb.Close();

            //GC.Collect();
            //GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(wb);
            //Marshal.FinalReleaseComObject(wb);
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
            log.Debug($"Open workbook in NEW excel application, please monitor it anything is wrong : {excelFileFullPath}");
            Excel.Workbook wb = null;
            Excel.Application xlApp = new Excel.Application();
            xlApp.Visible = true;
            wb = xlApp.Workbooks.Open(excelFileFullPath, UpdateLinks: false, ReadOnly: readOnly);

            xlApp.WindowState = Excel.XlWindowState.xlMaximized;
            xlApp.ActiveWindow.WindowState = Excel.XlWindowState.xlMaximized;
            if (!readOnly && wb.ReadOnly)
                MessageBox.Show("The excel file is opened as read only, but program is intended to write, there must be something wrong, please contact with IT support.", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return wb;
        }

        private static Excel.Workbook OpenExcelFileWithDefaultExcelApplication(string excelFileFullPath, bool readOnly = true)
        {
            Excel.Workbook wb = null;
            Excel.Application xlApp = GetDefaultFirstApplication();
            xlApp.Visible = true;
            wb = xlApp.Workbooks.Open(excelFileFullPath, UpdateLinks: false, ReadOnly: readOnly);

            xlApp.WindowState = Excel.XlWindowState.xlMaximized;
            xlApp.ActiveWindow.WindowState = Excel.XlWindowState.xlMaximized;

            if (!readOnly && wb.ReadOnly)
                MessageBox.Show("The excel file is opened as read only, but program is intended to write, there must be something wrong, please contact with IT support.", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return wb;
        }

        public static bool ExcelFileIsOpenedInEditMode(string excelFileFullPath)
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


        public static EnumExcelFileIsOpen ExcelFileIsOpened(string excelFileFullPath)
        {
            EnumExcelFileIsOpen res = EnumExcelFileIsOpen.FileIsNotOpen;

            string fileBaseName = Path.GetFileName(excelFileFullPath).ToUpper();

            foreach (Excel.Application item in GetAllExcelApplicationInstances())
            {
                foreach (Excel.Workbook eachWb in item.Workbooks)
                {
                    if (eachWb.FullName.ToUpper().Equals(excelFileFullPath.ToUpper()))
                    {
                        if (eachWb.ReadOnly)
                        {
                            if (!eachWb.Saved)
                            {
                                log.Debug($"workbook was opened readonly, with changes made.{eachWb.FullName}");
                                return EnumExcelFileIsOpen.FileIsOpen;
                            }
                            else
                            {
                                log.Debug($"workbook was opened readonly, without any change.{eachWb.FullName}");
                                return EnumExcelFileIsOpen.FileIsOpen;
                            }
                        }
                        else
                        {
                            if (!eachWb.Saved)
                            {
                                log.Debug($"workbook was opened in edit mode, with changes made.{eachWb.FullName}");
                                return EnumExcelFileIsOpen.FileIsOpen;
                            }
                            else
                            {
                                log.Debug($"workbook was opened in edit mode, without any change.{eachWb.FullName}");
                                return EnumExcelFileIsOpen.FileIsOpen;
                            }
                        }
                    }
                    else if (eachWb.Name.ToUpper().Equals(fileBaseName))
                    {
                        log.Debug($"same workbook name was found : {eachWb.FullName}");
                        //return ExcelFileOpenStatus.AnotherSameFileNameOpened;
                        res = res | EnumExcelFileIsOpen.AnotherSameFileNameOpened;
                    }
                }
            }

            if ((res & EnumExcelFileIsOpen.AnotherSameFileNameOpened) != 0)
                return EnumExcelFileIsOpen.AnotherSameFileNameOpened;

            return res;
        }
        public static ExcelFileOpenStatus ExcelFileIsOpenedAllStatus(string excelFileFullPath)
        {
            ExcelFileOpenStatus res = ExcelFileOpenStatus.None;

            string fileBaseName = Path.GetFileName(excelFileFullPath).ToUpper();

            foreach (Excel.Application item in GetAllExcelApplicationInstances())
            {
                //log.Debug($"Excel.Application, item.Workbooks:  {item.Workbooks.Count}");
                foreach (Excel.Workbook eachWb in item.Workbooks)
                {
                    //log.Debug($"eachWb: {eachWb.FullName}");
                    if (eachWb.FullName.ToUpper().Equals(excelFileFullPath.ToUpper()))
                    {
                        if (eachWb.ReadOnly)
                        {
                            if (!eachWb.Saved)
                            {
                                log.Debug($"workbook was opened readonly, with changes made.{eachWb.FullName}");
                                //res = res | ExcelFileOpenStatus.ReadOnlyButChangeNotSave;
                                res.SetFlag(ExcelFileOpenStatus.ReadOnlyButChangeNotSave);
                            }
                            else
                            {
                                log.Debug($"workbook was opened readonly, without any change.{eachWb.FullName}");
                                //res = res | ExcelFileOpenStatus.ReadOnlyAndNoChange;
                                res.SetFlag(ExcelFileOpenStatus.ReadOnlyAndNoChange);
                            }
                        }
                        else
                        {
                            if (!eachWb.Saved)
                            {
                                log.Debug($"workbook was opened in edit mode, with changes made.{eachWb.FullName}");
                                //res = res | ExcelFileOpenStatus.WriteButChangeNotSave;
                                res.SetFlag(ExcelFileOpenStatus.WriteButChangeNotSave);
                            }
                            else
                            {
                                log.Debug($"workbook was opened in edit mode, without any change.{eachWb.FullName}");
                                //res = res | ExcelFileOpenStatus.WriteAndNoChange;
                                res.SetFlag(ExcelFileOpenStatus.WriteAndNoChange);
                            }
                        }
                    }
                    else if (eachWb.Name.ToUpper().Equals(fileBaseName))
                    {
                        log.Debug($"same workbook name was found : {eachWb.FullName}");
                        //res = res | ExcelFileOpenStatus.AnotherSameFileNameOpened;
                        res.SetFlag(ExcelFileOpenStatus.AnotherSameFileNameOpened);
                    }
                }
            }

            return res;
        }

        public static Excel.Workbook CreateWorkbook(string excelFileFullPath, string sheetName = null, bool promptToOverwrite = true)
        {
            Excel.Workbook wb = null;

            if (File.Exists(excelFileFullPath))
            {
                if (promptToOverwrite)
                {
                    DialogResult response = MessageBox.Show(text: "文件已经存在，您要替换它吗？" + Environment.NewLine + excelFileFullPath
                        , caption: "File already exists", buttons: MessageBoxButtons.YesNoCancel, icon: MessageBoxIcon.Question);
                    if (response != DialogResult.Yes)
                        throw new OperationCanceledException();
                }

                var openStatus = ExcelFileIsOpened(excelFileFullPath);
                if (openStatus.HasFlag(EnumExcelFileIsOpen.FileIsOpen))
                {
                    MessageBox.Show("文件已打开，请先关闭它。", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    throw new OperationCanceledException();
                }
                else if (openStatus.HasFlag(EnumExcelFileIsOpen.AnotherSameFileNameOpened))
                {
                    MessageBox.Show("其他同名文件已打开，最好请先关闭它。", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    throw new OperationCanceledException();
                }
                //File.Delete(excelFileFullPath);
            }

            Excel.Application xlApp = GetDefaultFirstApplication();
            xlApp.SheetsInNewWorkbook = 1;
            wb = xlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            xlApp.ActiveWindow.DisplayGridlines = false;

            wb.Worksheets[1].Name = sheetName ?? Path.GetFileNameWithoutExtension(excelFileFullPath);

            Excel.XlFileFormat fileFormat = GetExcelFileFormatByFileExtension(excelFileFullPath);

            var orig = wb.Application.DisplayAlerts;
            wb.Application.DisplayAlerts = false;
            wb.SaveAs(Filename: excelFileFullPath, FileFormat: fileFormat);
            wb.Application.DisplayAlerts = orig;
            return wb;
        }
        public static Excel.Workbook GetActiveWorkbook()
        {
            return GetDefaultFirstApplication()?.ActiveWorkbook;
        }


        //public static Excel.Workbook GetExactWorkbook(string excelFileFullPath, out ExcelFileOpenStatus fileStatus, bool forEditPurpose = false)
        //{
        //    fileStatus = ExcelFileOpenStatus.FileNotOpen;

        //    Excel.Workbook wb = null;
        //    bool anotherSameBaseFileNameOpened = false;
        //    Excel.Workbook readOnlyNoChange = null;
        //    Excel.Workbook readOnlyChanged = null;
        //    Excel.Workbook writeNoChange = null;
        //    Excel.Workbook writeChanged = null;

        //    string fileBaseName = Path.GetFileName(excelFileFullPath).ToUpper();

        //    foreach (Excel.Application item in GetAllExcelApplicationInstances())
        //    {
        //        foreach (Excel.Workbook eachWb in item.Workbooks)
        //        {
        //            if (eachWb.FullName.ToUpper().Equals(excelFileFullPath.ToUpper()))
        //            {
        //                if (eachWb.ReadOnly)
        //                {
        //                    if (!eachWb.Saved)
        //                    {
        //                        readOnlyChanged = eachWb;
        //                        break;
        //                    }
        //                    else
        //                    {
        //                        readOnlyNoChange = eachWb;
        //                        break;
        //                    }
        //                }
        //                else
        //                {
        //                    if (!eachWb.Saved)
        //                    {
        //                        writeChanged = eachWb;
        //                        break;
        //                    }
        //                    else
        //                    {
        //                        writeNoChange = eachWb;
        //                        break;
        //                    }
        //                }
        //            }
        //            else if (eachWb.Name.ToUpper().Equals(fileBaseName))
        //            {
        //                wb = null;
        //                anotherSameBaseFileNameOpened = true;
        //                break;
        //            }
        //        }

        //        if (wb != null)
        //            break;
        //    }

        //    if (!forEditPurpose)
        //    {
        //        if (readOnlyNoChange != null)
        //        {
        //            return readOnlyNoChange;
        //        }
        //        else if (writeNoChange != null)
        //        {
        //            return writeNoChange;
        //        }
        //        else if (readOnlyChanged != null)
        //            else if (writeChanged != null)
        //            else
        //    }
        //    else
        //    {

        //    }
        //}


        //private static bool WorkbookExistsInExcelApplication(string excelFileName, out Excel.Workbook wb, Excel.Application xlApp = null)
        //{
        //    string fileBaseName = Path.GetFileName(excelFileName).ToUpper();
        //    wb = null;

        //    if (xlApp == null)
        //        xlApp = _ExcelApp ?? GetDefaultFirstApplication();

        //    foreach (Excel.Workbook eachWb in xlApp.Workbooks)
        //    {
        //        if (eachWb.Name.ToUpper().Equals(fileBaseName))
        //        {
        //            wb = eachWb;
        //            return true;
        //        }
        //    }

        //    return false;
        //}
        //private static bool WorkbookExistsInExcelApplication(string excelFileName, Excel.Application xlApp = null)
        //{
        //    Excel.Workbook wb = null;
        //    return WorkbookExistsInExcelApplication(excelFileName, out wb, xlApp);
        //}


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
                    throw new ArgumentException("File extension is invalid :" + ext.ToLower());
                    //break;
            }
            return fileFormat;
        }
        #endregion

        #region sheet name
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

        private static Excel.Workbook SetParamWorkbookToAcitveWorkbookWhenItIsNull(Excel.Workbook wb)
        {
            if (wb != null)
            {
                return wb;
            }
            else
            {
                if (((Excel.Application)wb.Application).Workbooks.Count > 0)
                    return ((Excel.Application)wb.Application).ActiveWorkbook;
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
            //Stopwatch stopwatch = Stopwatch.StartNew();
            //log.Debug($"(Performance) GetAllExcelApplicationInstances started ...: {stopwatch.Elapsed}");

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
                //log.Debug($"appName: {appName}");
                if (appName == "Microsoft Excel")
                {
                    Excel.Application xlApp = wb.Application;

                    //log.Debug($"wb : {wb.Name}, readonly: {wb.ReadOnly}");

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
            //stopwatch.Stop();
            //log.Debug($"(Performance) GetAllExcelApplicationInstances (total) time consumed: {stopwatch.Elapsed}");
            //return result.ToList();
        }
        // Requires Using System.Runtime.InteropServices.ComTypes

        // Get all running instance by querying ROT

        //public static List<object> GetRunningInstances(string[] progIds)
        //{
        //    List<string> clsIds = new List<string>();

        //    // get the app clsid
        //    foreach (string progId in progIds)
        //    {
        //        Type type = Type.GetTypeFromProgID(progId);
        //        if (type != null)
        //            clsIds.Add(type.GUID.ToString().ToUpper());
        //    }

        //    //Guid excelAppClsId = new Guid("{00000000-0000-0000-C000-000000000046}");
        //    //clsIds.Add(excelAppClsId.ToString().ToUpper());

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

        //        log.Debug($"displayName {displayName}");


        //        foreach (string clsId in clsIds)
        //        {
        //            if (displayName.ToUpper().IndexOf(clsId) > 0)
        //            {
        //                object ComObject;
        //                Rot.GetObject(monikers[0], out ComObject);

        //                if (ComObject == null)
        //                    continue;
        //                instances.Add(ComObject);
        //                break;
        //            }
        //        }
        //    }
        //    return instances;
        //}

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
        //private static Excel.Application GetExcelApplication(bool MultiInstanceThenGetTheDefault = false)
        //{
        //    if (_ExcelApp != null)
        //        return _ExcelApp;

        //    Excel.Application xlApp1 = null;
        //    int cnt = 0;

        //    foreach (Excel.Application xlApp2 in GetAllExcelApplicationInstances())
        //    {
        //        if (MultiInstanceThenGetTheDefault)
        //        {
        //            xlApp1 = xlApp2;
        //            xlApp1.Visible = true;
        //            break;
        //        }
        //        else
        //        {
        //            cnt++;
        //            if (cnt > 1) break;
        //        }
        //    }
        //    if (!MultiInstanceThenGetTheDefault && cnt > 1)
        //        throw new Exception("Multiple Excel instances are running, please change your program to get the excel application directly " + Environment.NewLine + "Like moniker.bindtomonier(filename), or set this _xlapp");

        //    if (xlApp1 == null)
        //    {
        //        try
        //        {
        //            xlApp1 = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
        //        }
        //        catch (Exception)
        //        {
        //            xlApp1 = new Excel.Application();
        //        }
        //        finally
        //        {
        //            xlApp1.Visible = true;
        //        }
        //    }

        //    return xlApp1;

        //    //Excel.Application xlApp = GetAllExcelApplicationInstances().FirstOrDefault();

        //    //if (xlApp ==null)
        //    //{
        //    //    xlApp = GetDefaultFirstApplication();
        //    //    xlApp.Visible = true;
        //    //}
        //    //else
        //    //{
        //    //    if (MultiInstanceThenGetTheDefault)
        //    //    {
        //    //        xlApp = GetDefaultFirstApplication();
        //    //        xlApp.Visible = true;
        //    //    }
        //    //    else
        //    //    {
        //    //        throw new Exception("Multiple Excel instances are running, please change your program to get the excel application directly " + Environment.NewLine + "Like moniker.bindtomonier(filename), or set this _xlapp");
        //    //    }
        //    //}
        //    //return xlApp;
        //}
        #endregion

        #region Excel sheet names
        public static IEnumerable<string> GetAllSheetNames(string excelFileFullPath)
        {
            //bool useTra = false;

            if (!File.Exists(excelFileFullPath))
            {
                //MessageBox.Show($@"文件不存在：{excelFileFullPath}");
                throw new System.IO.FileNotFoundException($@"文件不存在：{excelFileFullPath}");
            }

            IEnumerable<string> res = null;
            try
            {
                log.Debug($"Get all sheet names of : {excelFileFullPath}");
                res = ExcelUtilOX.GetAllSheetNamesViaOpenXml(excelFileFullPath);
                log.Debug($"GetAllSheetNamesViaOpenXml: {string.Join(",", res.ToArray())}");
            }
            catch (FileFormatException ex)
            {
                log.Debug("The file is probably not Excel 2007 later. so program has to use tradition way.");
                //log.Debug(ex.ToString());
                return GetAllSheetNamesWithOleDb(excelFileFullPath);
                // use traditional way 
                //useTra = true; 
            }
            catch (IOException ex)
            {
                //file is already open
                log.Debug("The file is Excel 2007, or it is open, so program has to use tradition way.");
                log.Debug(ex.ToString());

                if (Path.GetExtension(excelFileFullPath).Equals(".xls", StringComparison.OrdinalIgnoreCase))
                {
                    return GetAllSheetNamesWithOleDb(excelFileFullPath);
                }
                // use traditional way 
                //useTra = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }

            if (res == null || res.Count() <= 0)
            {
                bool alreadyOpened;
                Excel.Workbook wb = OpenWorkbook(excelFileFullPath, out alreadyOpened, true);
                res = wb.Worksheets.Cast<Excel.Worksheet>().Select(a => a.Name).ToList();

                if (!alreadyOpened)
                    wb.CloseWithoutSave();
            }
            return res;
        }


        public static IEnumerable<string> GetAllSheetNamesWithOleDb(string excelFile)
        {
            System.Data.DataTable dt = null;
            try
            {
                string connString = null;
                string extension = Path.GetExtension(excelFile);
                switch (extension)
                {
                    case ".xls": //Excel 97-03.
                        //connString = $@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={excelFile};Extended Properties=Excel 8.0;";
                        connString = $@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={excelFile};Extended Properties='Excel 8.0;HDR=YES;IMEX=1'";
                        break;
                    case ".xlsm": //Excel 07 or higher.
                    case ".xlsx": //Excel 07 or higher.
                        connString = $@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = { excelFile }; Extended Properties ='Excel 8.0;HDR=YES;IMEX=1'";
                        break;
                }

                using (OleDbConnection objConn = new OleDbConnection(connString))
                {
                    objConn.Open();
                    dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    objConn.Close();
                }

                if (dt == null)
                    return null;

                HashSet<string> sheets = new HashSet<string>();

                dt.Rows.Cast<DataRow>().AsEnumerable().Select(a => a["TABLE_NAME"].ToString()).ToList().ForEach(b =>
                {
                    string shtName = b.Split('$')[0].Replace("'", "");
                    if (!sheets.Contains(shtName))
                        sheets.Add(shtName);
                });

                //String[] excelSheets = new String[dt.Rows.Count];
                //int i = 0;
                //foreach (System.Data.DataRow row in dt.Rows)
                //{
                //    excelSheets[i] = row["TABLE_NAME"].ToString();
                //    i++;
                //}

                log.Debug($"Get all sheets name via oledb: {string.Join(",", sheets.ToArray())}");
                return sheets.AsEnumerable();
            }
            catch (Exception ex)
            {
                throw ex;
                return null;
            }
        }

        public static List<string> GetAllSheetNamesWithOleDb2(string filePath)
        {
            OleDbConnectionStringBuilder sbConnection = new OleDbConnectionStringBuilder();
            String strExtendedProperties = String.Empty;
            sbConnection.DataSource = filePath;
            if (Path.GetExtension(filePath).Equals(".xls"))//for 97-03 Excel file
            {
                sbConnection.Provider = "Microsoft.Jet.OLEDB.4.0";
                strExtendedProperties = "Excel 8.0;HDR=Yes;IMEX=1";//HDR=ColumnHeader,IMEX=InterMixed
            }
            else if (Path.GetExtension(filePath).Equals(".xlsx"))  //for 2007 Excel file
            {
                sbConnection.Provider = "Microsoft.ACE.OLEDB.12.0";
                strExtendedProperties = "Excel 12.0;HDR=Yes;IMEX=1";
            }
            sbConnection.Add("Extended Properties", strExtendedProperties);
            List<string> listSheet = new List<string>();
            using (OleDbConnection conn = new OleDbConnection(sbConnection.ToString()))
            {
                conn.Open();
                DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                foreach (DataRow drSheet in dtSheet.Rows)
                {
                    if (drSheet["TABLE_NAME"].ToString().Contains("$"))//checks whether row contains '_xlnm#_FilterDatabase' or sheet name(i.e. sheet name always ends with $ sign)
                    {
                        listSheet.Add(drSheet["TABLE_NAME"].ToString());
                    }
                }
            }
            return listSheet;
        }
        #endregion

        #region read sheet data
        public static List<T> ReadData<T>(this Excel.Worksheet sht, long dataFromRow = 2, long fromCol = 1, long toCol = 0)
        {
            List<T> result = null;

            dynamic[,] arrData = sht.ReadData(dataFromRow, fromCol, toCol);

            if (arrData == null) return result;
            if (arrData.Rank != 2) throw new InvalidOperationException("arrData.Rank != 2 ReadData");
            if (arrData.GetUpperBound(0) < arrData.GetLowerBound(0)) return result;

            Dictionary<PropertyInfo, int> propColNum = new Dictionary<PropertyInfo, int>();
            Dictionary<PropertyInfo, string> propColHeader = new Dictionary<PropertyInfo, string>();
            Dictionary<PropertyInfo, string> propColFormat = null;

            PropertyInfo[] props = typeof(T).GetProperties(BindingFlags.Instance | BindingFlags.Public);
            foreach (var prop in props)
            {
                SheetColumnIndexAttribute colNumAtt = prop.GetCustomAttribute<SheetColumnIndexAttribute>();
                if (colNumAtt != null) propColNum.Add(prop, colNumAtt.ColumnNum);

                SheetColumnHeaderAttribute colHeaderAtt = prop.GetCustomAttribute<SheetColumnHeaderAttribute>();
                if (colHeaderAtt != null) propColHeader.Add(prop, colHeaderAtt.ColumnHeader);

                SheetColumnFormatAttribute colFormatAtt = prop.GetCustomAttribute<SheetColumnFormatAttribute>();
                if (colFormatAtt != null)
                {
                    if (propColFormat == null) propColFormat = new Dictionary<PropertyInfo, string>();
                    propColFormat.Add(prop, colFormatAtt.ColumnFormat);
                }
            }

            Excel.Range header = sht.GetRange(dataFromRow - 1, 1, dataFromRow - 1);
            foreach (KeyValuePair<PropertyInfo, string> item in propColHeader)
            {
                PropertyInfo prop = item.Key;
                string colHeader = item.Value;

                Excel.Range rgFound = null;
                int colHeaderFountCnt = header.FindInRange(colHeader, out rgFound);
                if (colHeaderFountCnt == 0) throw new InvalidOperationException($"You specified column header {colHeader} for property {prop} in class {typeof(T).Name}, but it cannot be found from the header.");
                if (colHeaderFountCnt > 1) throw new InvalidOperationException($"You specified column header {colHeader} for property {prop} in class {typeof(T).Name}, but {colHeaderFountCnt} same header were found from the header.");

                int foundColNum = rgFound.Column;
                if (propColNum.ContainsKey(prop))
                {
                    if (propColNum[prop] != foundColNum) throw new InvalidOperationException($"You specified both column header {colHeader} and column number {propColNum[prop]} for property {prop} in class {typeof(T).Name}, but the column number detected by column header is {foundColNum}");
                }
                else
                    propColNum.Add(prop, foundColNum);
            }
            result = Converter.ConvertArrayToList<T>(arrData, propColNum, propColFormat);
            return result;
        }
        public static dynamic[,] ReadData(this Excel.Worksheet sht, long dataFromRow = 2, long fromCol = 1, long toCol = 0)
        {
            return sht.GetRange(dataFromRow, fromCol, 0, toCol).ReadData();
        }
        public static dynamic[,] ReadData(this Excel.Range rg)
        {
            dynamic[,] result = null;

            if (rg != null)
            {
                if (rg.Cells.Count == 1)
                {
                    Array tmp = Array.CreateInstance(typeof(object), new int[] { 1, 1 }, new int[] { 1, 1 });
                    tmp.SetValue(rg.Value, 1, 1);
                    result = (dynamic[,])tmp;
                }
                else
                    result = rg.Value;
            }

            return result;
        }
        public static Excel.Range GetRange(this Excel.Worksheet sht, long StartRow = 1, long StartCol = 1, long EndRow = 0, long EndCol = 0)
        {
            if (EndRow <= 0)
                EndRow = sht.MaxRow();
            if (EndCol <= 0)
                EndCol = sht.MaxCol();

            if (StartRow > EndRow)
                throw new ArgumentException("StartRow > EndRow to method GetRange");
            if (StartCol > EndCol)
                throw new ArgumentException("StartCol > EndCol to method GetRange");

            return sht.Range[sht.Cells[StartRow, StartCol], sht.Cells[EndRow, EndCol]];
        }
        public static long MaxRow(this Excel.Worksheet sht)
        {
            long result = 0;

            Excel.Range lastRow = sht.Cells.Find("*", After: sht.Cells[1, 1], SearchOrder: Excel.XlSearchOrder.xlByRows, SearchDirection: Excel.XlSearchDirection.xlPrevious);
            result = lastRow == null ? 0 : lastRow.Row;

            if (result == 1)
            {
                Excel.Range lastCol = ((Excel.Range)sht.Cells[1, sht.Columns.Count]).End[Excel.XlDirection.xlToLeft];

                if (lastCol.Column <= 0)
                    log.Error($"sheet 's End[Excel.XlDirection.xlToLeft].Column <= 0");
                else if (lastCol.Column == 1)
                    result = lastCol.Value == null || string.IsNullOrWhiteSpace((string)lastCol.Value) ? 0 : 1;
                else
                {
                    var usedCellNum = sht.Application.WorksheetFunction.CountA(sht.Range[sht.Cells[1, 1], lastCol]);
                    if (usedCellNum == 1)
                        result = lastCol.Value == null || string.IsNullOrWhiteSpace((string)lastCol.Value) ? 0 : 1;
                    else if (usedCellNum <= 0)
                        result = 0;
                    else
                        result = 1;
                }
            }

            return result;
        }
        public static long MaxCol(this Excel.Worksheet sht)
        {
            long result = 0;

            Excel.Range lastColumn = sht.Cells.Find("*", After: sht.Cells[1, 1], SearchOrder: Excel.XlSearchOrder.xlByColumns, SearchDirection: Excel.XlSearchDirection.xlPrevious);
            result = lastColumn == null ? 0 : lastColumn.Column;

            if (result == 1)
            {
                Excel.Range lastRow = ((Excel.Range)sht.Cells[sht.Rows.Count, 1]).End[Excel.XlDirection.xlUp];

                if (lastRow.Row <= 0)
                    log.Error($"sheet 's End[Excel.XlDirection.xlUp].Column <= 0");
                else if (lastRow.Row == 1)
                    result = lastRow.Value == null || string.IsNullOrWhiteSpace((string)lastRow.Value) ? 0 : 1;
                else
                {
                    var usedCellNum = sht.Application.WorksheetFunction.CountA(sht.Range[sht.Cells[1, 1], lastRow]);
                    if (usedCellNum == 1)
                        result = lastRow.Value == null || string.IsNullOrWhiteSpace((string)lastRow.Value) ? 0 : 1;
                    else if (usedCellNum <= 0)
                        result = 0;
                    else
                        result = 1;
                }
            }

            return result;
        }
        public static int FindInRange(this Excel.Range rgFindIn, string whatToFind, out Excel.Range rgFound)
        {
            int foundNum = 0;
            rgFound = FindAllInRange(rgFindIn, whatToFind, out foundNum, false);
            return foundNum;
        }
        public static Excel.Range FindAllInRange(this Excel.Range rgFindIn, string whatToFind, out int foundNum, bool findFirstThenStop = false)
        {
            Excel.Range rgResult = null;
            Excel.Range rgFound = null;
            string firstAddress = string.Empty;
            foundNum = 0;

            rgFound = rgFindIn.Find(What: whatToFind, After: rgFindIn.Cells[rgFindIn.Rows.Count, rgFindIn.Columns.Count]
                , LookIn: Excel.XlFindLookIn.xlValues, LookAt: Excel.XlLookAt.xlWhole, SearchOrder: Excel.XlSearchOrder.xlByRows, MatchCase: false, MatchByte: false, SearchFormat: false);

            if (!findFirstThenStop)
            {
                if (rgFound != null)
                {
                    firstAddress = rgFound.Address;
                    foundNum++;

                    if (foundNum == 1) rgResult = rgFound;

                    while (true)
                    {
                        rgFound = rgFindIn.Find(What: whatToFind, After: rgFound
                            , LookIn: Excel.XlFindLookIn.xlValues, LookAt: Excel.XlLookAt.xlWhole, SearchOrder: Excel.XlSearchOrder.xlByRows, MatchCase: false, MatchByte: false, SearchFormat: false);
                        if (rgFound == null) break;
                        if (rgFound.Address.Equals(firstAddress)) break;
                        foundNum++;
                    }
                }
            }
            else
                rgResult = rgFound;

            return rgResult;
        }
        public static void AppendData<T>(this Excel.Worksheet sht, IEnumerable<T> data, long startRow = 0, int startCol = 1, bool writeHeader = false)
        {
            long fromRow = startRow == 0 ? sht.MaxRow() + 1 : startRow;

            if (writeHeader)
            {//todo
                PropertyInfo[] props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

                int totalPropsNum = 0;
                foreach (PropertyInfo prop in props)
                {
                    var att = prop.GetCustomAttribute<DoNotExportToWorksheetAttribute>();
                    if (att == null)
                        totalPropsNum++;
                }

                var arrHeader = Array.CreateInstance(typeof(object), new int[] { 1, totalPropsNum }, new int[] { 1, 1 });
                totalPropsNum = 0;
                foreach (PropertyInfo prop in props)
                {
                    var att = prop.GetCustomAttribute<DoNotExportToWorksheetAttribute>();
                    if (att == null)
                    {
                        totalPropsNum++;
                        arrHeader.SetValue(prop.Name, 1, totalPropsNum);
                    }
                }

                ((Excel.Range)sht.Cells[fromRow, 1]).Resize[1, totalPropsNum].Value = arrHeader;
                fromRow++;
            }

            ADODB.Recordset rs = Converter.ConvertListToRecordSet(data);
            ((Excel.Range)sht.Cells[fromRow, 1]).CopyFromRecordset(rs);
            rs.Close();
            Marshal.ReleaseComObject(rs);
        }
        public static void SetConditionalFormatForBorder(this Excel.Worksheet sht, long fromRow = 0, long toRow = 0, int toColumn = 0, object arrKeyColumnsNotBlank = null)
        {
            try
            {
                long _startRow = fromRow == 0 ? 2 : fromRow;
                long _endRow = toRow == 0 ? sht.MaxRow() : toRow;
                long _endCol = toColumn == 0 ? sht.MaxCol() : toColumn;

                string formula = string.Empty;
                if (arrKeyColumnsNotBlank != null)
                {
                    _endRow = _endRow + 10000;
                    if (arrKeyColumnsNotBlank.GetType().IsArray)
                    {
                        StringBuilder formulaStr = new StringBuilder();

                        object[] arrKeyCols = (object[])arrKeyColumnsNotBlank;
                        for (int i = arrKeyCols.GetLowerBound(0); i <= arrKeyCols.GetUpperBound(0); i++)
                        {
                            formulaStr.Append($",len(trim(${Converter.Num2Letter(Convert.ToInt32(arrKeyCols[i]))}{_startRow}))");
                        }
                        formula = formulaStr.ToString();
                        if (formula.Length > 0)
                        {
                            formula = formula.Substring(1, formula.Length - 1);
                            formula = $"=And({formula})";
                        }
                    }
                    else
                    {
                        formula = $"=len(trim(${Converter.Num2Letter(Convert.ToInt32(arrKeyColumnsNotBlank))}{_startRow})) > 0";
                    }
                }
                if (formula.Length <= 0) formula = "=1=1";
                if (_endRow < _startRow) return;

                Excel.Range rgToFormat = sht.GetRange(_startRow, 1, _endRow, _endCol);
                Excel.FormatCondition formatCondition = (Excel.FormatCondition)rgToFormat.FormatConditions.Add(Type: Excel.XlFormatConditionType.xlExpression, Formula1: formula);
                formatCondition.SetFirstPriority();
                formatCondition.StopIfTrue = false;
                formatCondition.Borders.Weight = Excel.XlBorderWeight.xlHairline;
            }
            catch (Exception ex)
            {
                log.Debug($"SetConditionalFormatForBorder got error for the sheet {sht.Name} of workbook {((Excel.Workbook)sht.Parent).Name}  ");
                log.Error(ex.Message + Environment.NewLine + ex.InnerException?.Message + Environment.NewLine + ex.ToString());
            }
        }
        public static void SetConditionalFormatForOddEvenLine(this Excel.Worksheet sht, long fromRow = 0, long toRow = 0, int toColumn = 0, object arrKeyColumnsNotBlank = null)
        {
            try
            {
                long _startRow = fromRow == 0 ? 2 : fromRow;
                long _endRow = toRow == 0 ? sht.MaxRow() : toRow;
                long _endCol = toColumn == 0 ? sht.MaxCol() : toColumn;

                string formula = string.Empty;
                if (arrKeyColumnsNotBlank != null)
                {
                    _endRow = _endRow + 10000;
                    if (arrKeyColumnsNotBlank.GetType().IsArray)
                    {
                        StringBuilder formulaStr = new StringBuilder();

                        object[] arrKeyCols = (object[])arrKeyColumnsNotBlank;
                        for (int i = arrKeyCols.GetLowerBound(0); i <= arrKeyCols.GetUpperBound(0); i++)
                        {
                            formulaStr.Append($",len(trim(${Converter.Num2Letter(Convert.ToInt32(arrKeyCols[i]))}{_startRow}))");
                        }
                        formula = formulaStr.ToString();
                        if (formula.Length > 0)
                        {
                            formula = formula.Substring(1, formula.Length - 1);
                            formula = $"=And({formula})";
                        }
                    }
                    else
                    {
                        formula = $"=len(trim(${Converter.Num2Letter(Convert.ToInt32(arrKeyColumnsNotBlank))}{_startRow})) > 0";
                    }
                }

                if (_endRow < _startRow) return;

                string formulaOdd = string.Empty;
                string formulaEven = string.Empty;
                if (formula.Length > 0)
                {
                    formula = formula.Substring(1, formula.Length - 1);

                    formulaOdd = $"=And({formula},mod(row(),2)=0)";
                    formulaEven = $"=And({formula},mod(row(),2)<>0)";
                }
                else
                {
                    formulaOdd = $"=mod(row(),2)=0";
                    formulaEven = $"=mod(row(),2)<>0";
                }

                Excel.Range rgToFormat = sht.GetRange(_startRow, 1, _endRow, _endCol);
                Excel.FormatCondition formatCondition = (Excel.FormatCondition)rgToFormat.FormatConditions.Add(Type: Excel.XlFormatConditionType.xlExpression, Formula1: formulaOdd);
                formatCondition.SetFirstPriority();
                formatCondition.StopIfTrue = false;
                formatCondition.Interior.Color = 14348258;

                formatCondition = (Excel.FormatCondition)rgToFormat.FormatConditions.Add(Type: Excel.XlFormatConditionType.xlExpression, Formula1: formulaEven);
                formatCondition.SetFirstPriority();
                formatCondition.StopIfTrue = false;
                formatCondition.Interior.Color = 14083324;
            }
            catch (Exception ex)
            {
                log.Debug($"SetConditionalFormatForOddEvenLine got error for the sheet {sht.Name} of workbook {((Excel.Workbook)sht.Parent).Name}  ");
                log.Error(ex.Message + Environment.NewLine + ex.InnerException?.Message + Environment.NewLine + ex.ToString());
            }
        }
        public static void FreezeSheet(this Excel.Worksheet sht, int splitColumn = 0, long splitRow = 1)
        {
            try
            {
                var orig = sht.Visible;
                sht.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                sht.Activate();
                var xlApp = sht.Application;

                xlApp.ActiveWindow.FreezePanes = false;
                xlApp.ActiveWindow.SplitColumn = splitColumn;
                xlApp.ActiveWindow.SplitRow = (int)splitRow;
                xlApp.ActiveWindow.FreezePanes = true;

                sht.Visible = orig;
            }
            catch (Exception ex)
            {
                log.Debug($"FreezeSheet got error, the sheet {sht.Name} of workbook {((Excel.Workbook)sht.Parent).Name} may be not frozen ");
                log.Error(ex.Message + Environment.NewLine + ex.InnerException?.Message + Environment.NewLine + ex.ToString());
            }
        }
        public static void FormatHeader(this Excel.Worksheet sht, long headerAtRow = 1)
        {
            Excel.Range header = sht.GetRange(headerAtRow, 1, headerAtRow);

            header.Borders[Excel.XlBordersIndex.xlDiagonalDown].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            header.Borders[Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Excel.XlLineStyle.xlLineStyleNone;

            header.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            header.Borders.Weight = Excel.XlBorderWeight.xlMedium;
            header.Font.Bold = true;
            header.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            header.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGreen);
        }
        public static void AddAutoFilter(this Excel.Worksheet sht, long filterAtRow = 1, long toColumn = 0)
        {
            long maxCol = toColumn == 0 ? sht.MaxCol() : toColumn;
            Excel.Range header = sht.GetRange(filterAtRow, 1, filterAtRow, maxCol);
            if (sht.AutoFilterMode)
                sht.AutoFilterMode = false;

            header.AutoFilter(Field:header.Column);
        }

        #endregion

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        public static void BringToFront(this Excel.Application xlApp)
        {
            string caption = xlApp.Caption;
            IntPtr handler = FindWindow(null, caption);
            SetForegroundWindow(handler);
        }
        public static void BringToFront(this Excel.Workbook wb)
        {
            wb.Activate();
            wb.Application.ActiveWindow.Activate();
            BringToFront(wb.Application);
        }
    }
}
