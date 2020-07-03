using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommonLib
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class SheetColumnIndexAttribute : System.Attribute
    {
        private int _columnNum;
        public int ColumnNum
        {
            get { return _columnNum; }
        }

        private string _columnLetter;
        public string ColumnLetter
        {
            get { return _columnLetter; }
        }

        public SheetColumnIndexAttribute(int columnNum)
        {
            this._columnNum = columnNum;
            this._columnLetter = Converter.Num2Letter(columnNum);
        }

        public SheetColumnIndexAttribute(string columnLetter)
        {
            this._columnLetter = columnLetter;
            this._columnNum = Converter.Letter2Num(columnLetter);
        }
    }
}
