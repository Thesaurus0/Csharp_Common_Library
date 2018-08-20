using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CommonLib
{
    public class Class1
    {
        //// System.Runtime.InteropServices
        //[DllImport("ole32.dll")]
        //public static extern int GetRunningObjectTable(int reserved, out System.Runtime.InteropServices.UCOMIRunningObjectTable prot);
        //[DllImport("ole32.dll")]
        //public static extern int CreateBindCtx(int reserved, out System.Runtime.InteropServices.UCOMIBindCtx ppbc);
        //    private void btnDprocess_Click(object sender, EventArgs e)
        //    {
        //        List<object> list = new List<object>();
        //        int numFetched;
        //        UCOMIRunningObjectTable runningObjectTable;
        //        UCOMIEnumMoniker monikerEnumerator;
        //        UCOMIMoniker[] monikers = new UCOMIMoniker[1];
        //        GetRunningObjectTable(0, out runningObjectTable);
        //        runningObjectTable.EnumRunning(out monikerEnumerator);
        //        monikerEnumerator.Reset();
        //        while (monikerEnumerator.Next(1, monikers, out numFetched) == 0)
        //        {
        //            UCOMIBindCtx ctx;
        //            CreateBindCtx(0, out ctx);
        //            string runningObjectName;
        //            monikers[0].GetDisplayName(ctx, null, out runningObjectName);
        //            AInfo(runningObjectName);
        //            Guid g = new Guid();
        //            monikers[0].GetClassID(out g);
        //            AInfo(g.ToString());
        //            object runningObjectVal;
        //            runningObjectTable.GetObject(monikers[0], out runningObjectVal);
        //            list.Add(runningObjectVal);
        //        }
        //        for (int i = 0; i < list.Count; i++)
        //        {
        //            OfficeExcel._Application xls = list[i] as OfficeExcel._Application;
        //            if (xls == null)
        //                continue;
        //            try
        //            {
        //                this.listBox1.Items.Add(i.ToString("D3") + "\t" + xls.Workbooks[1].Name);
        //            }
        //            catch { }
        //        }
        //    }


        //public static void GetExcelpro2(ParamType1 ParamObj1, ParamType2 ParamObj2 = null)
        //{
        //    System.Diagnostics.Process.GetProcessesByName("")
        //}


        public static void GetAllPr()
        {
            string fileName = null;
            Process[] pl = Process.GetProcesses();
            for (int i = 0; i < pl.Length; i++)
            {
                try
                {
                    if (pl[i].ProcessName.ToLower() == "excel")
                    {
                        MessageBox.Show(pl[i].Id.ToString());
                        Microsoft.Office.Interop.Excel.Application ap = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                        MessageBox.Show(ap.Workbooks.Count.ToString());//问题表现

                        bool Invalid = true;
                        for (int j = 1; j <= ap.Workbooks.Count; j++)
                        {
                            Microsoft.Office.Interop.Excel.Workbook book = ap.Workbooks[j];
                            if (book == null)
                            {
                                //ap.Quit();
                                //pl[i].Kill();
                            }
                            else
                            {
                                Invalid = false;
                                //MessageBox.Show(book.Name);
                                fileName = ((Microsoft.Office.Interop.Excel.Worksheet)book.ActiveSheet).Name;
                                //book.SaveAs(filePath + fileName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                //Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                                //ap.Quit();
                                //pl[i].Kill();
                                break;
                            }
                        }

                        if (Invalid)
                        {
                            ap.Quit();
                            pl[i].Kill();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message + "\n" + ex.StackTrace + "\n" + ex.Source + "\n" + ex.TargetSite);
                }
            }
        }
    }
}
