using System;
using System.CodeDom;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CommonLib
{
    //enum enFileType
    //{ 
    //Excel_XLSX,

    //}
    public class Fundamentals
    {
        public static string SelectSingleFile(string fileExtenstionName, string fileExtenstion, string defaultFilePath = "", string title = "")
        {
            //Excel File|*.xlsx;*.xls;*.xlsm;*.xls*
            //Text files (*.txt)|*.txt

            string defaultDirectory = null;

            if (File.Exists(defaultFilePath))
            {
                defaultDirectory = System.IO.Path.GetDirectoryName(defaultFilePath);
            }
            else if (Directory.Exists(defaultFilePath))
            {
                defaultDirectory = defaultFilePath;
            }
            else
            {
                defaultDirectory = AppDomain.CurrentDomain.BaseDirectory;
            }

            OpenFileDialog fd = new OpenFileDialog();
            fd.InitialDirectory = defaultDirectory;
            fd.Title = string.IsNullOrWhiteSpace(title) ? "请选文件" : title;
            //fd.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
            fd.Filter = fileExtenstionName + "|" + fileExtenstion + "|All files|*.*";
            //fd.Filter
            fd.DefaultExt = "xlsx";
            fd.FilterIndex = 1;

            fd.Multiselect = false;
            //fd.ReadOnlyChecked = true;
            //fd.ShowReadOnly = true;
            //fd.CheckFileExists = true;
            //fd.CheckPathExists = true;
            DialogResult rs = fd.ShowDialog();
            if (rs == DialogResult.OK)
                return fd.FileName;
            else
                return string.Empty;
        }

        public static bool PathIsValidDirectory(string filePath)
        {
            FileAttributes attr = File.GetAttributes(filePath);
            return attr.HasFlag(FileAttributes.Directory);
        }
    }
}
