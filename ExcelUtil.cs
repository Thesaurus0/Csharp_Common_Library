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

        public static Excel.Worksheet GetSheetBySheetName(string shtName, Excel.Workbook wb = null)
        {
            Excel.Worksheet shtOut = null;



            return shtOut;
        }

         
        public static Excel.Range GetRangeByStartEndPos(Excel.Worksheet sht, long StartRow  , long StartCol, long EndRow, long EndCol )
        {
            Excel.Range rg = null;

            if (StartRow > EndRow)
                throw new ArgumentException("StartRow > EndRow");
            if (StartCol > EndCol)
                throw new ArgumentException("StartCol > EndCol");

            rg = sht.Range[sht.Cells[StartRow, StartCol], sht.Cells[EndRow, EndCol]];
            return rg;
        }

    }
}
