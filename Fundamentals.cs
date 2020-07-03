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
    [Flags]
    public enum ExcelFileOpenStatus
    {
        None = 0,
        //FileNotOpen = None,
        ReadOnlyAndNoChange = 2,
        ReadOnlyButChangeNotSave = 4,
        WriteAndNoChange = 8,
        WriteButChangeNotSave = 16,
        AnotherSameFileNameOpened = 32,
        //FileIsOpen = ReadOnlyAndNoChange | ReadOnlyButChangeNotSave  | WriteAndNoChange | WriteButChangeNotSave  
    }
    public enum EnumExcelFileIsOpen
    {
        //None = 0,
        FileIsNotOpen = 0,
        FileIsOpen = 1,
        AnotherSameFileNameOpened = 2
    }
    public static class Fundamentals
    {
        public static string SelectSaveAsFileDialog(string fileExtenstionName, string fileExtenstion, string defaultFilePath = "", string title = "")
        {
            //Excel File|*.xlsx
            //Text files (*.txt)|*.txt
            string defaultDirectory;
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
                string nearest = FindExistingNearestFolder(defaultFilePath);
                if (string.IsNullOrWhiteSpace(nearest) || nearest.Length <= 5)
                {
                    defaultDirectory = AppDomain.CurrentDomain.BaseDirectory;
                }
                else
                {
                    defaultDirectory = nearest;
                }
            }

            SaveFileDialog fd = new SaveFileDialog();
            fd.InitialDirectory = defaultDirectory;
            fd.Title = string.IsNullOrWhiteSpace(title) ? "保存文件" : title;
            //fd.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
            fd.Filter = fileExtenstionName + "|" + fileExtenstion + "|All files|*.*";
            fd.DefaultExt = "xlsx";
            fd.FilterIndex = 1;

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
                string nearest = FindExistingNearestFolder(defaultFilePath);
                if (string.IsNullOrWhiteSpace(nearest) || nearest.Length <= 5)
                {
                    defaultDirectory = AppDomain.CurrentDomain.BaseDirectory;
                }
                else
                {
                    defaultDirectory = nearest;
                }
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

        public static string FindExistingNearestFolder(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return string.Empty;

            string parentF = Path.GetDirectoryName(filePath);
            string rootPath = Path.GetPathRoot(filePath);

            while (parentF != rootPath)
            {
                if (File.Exists(parentF))
                {
                    return parentF;
                }
                else if (Directory.Exists(parentF))
                {
                    return parentF;
                }

                parentF = Path.GetDirectoryName(parentF);
            }

            return rootPath;
        }

        public static void SetFlag(this ref ExcelFileOpenStatus target , ExcelFileOpenStatus newTag) {
            target =  target | newTag;
        }

        public static void UnsetFlag(this ref ExcelFileOpenStatus target, ExcelFileOpenStatus tagToBeRemoved)
        {
            target = target & (~tagToBeRemoved);
        }

        // works with "None" as well
        public static bool HasFlagSecure(this ExcelFileOpenStatus target, ExcelFileOpenStatus tag)
        {
            return (target & tag ) == tag;
        }// works with "None" as well
        public static void ToggleFlag(this ref ExcelFileOpenStatus target, ExcelFileOpenStatus tag)
        {
            target = target ^ tag;
        }
    }
}
