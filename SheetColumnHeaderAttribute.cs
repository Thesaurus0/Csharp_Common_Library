using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommonLib
{
    [AttributeUsage( AttributeTargets.Property, AllowMultiple =false, Inherited = false)]
    public class SheetColumnHeaderAttribute : System.Attribute
    {
        private string _header;

        public string ColumnHeader {
            get { return _header; }
        }

        public SheetColumnHeaderAttribute(string header) {
            this._header = header;
        }
    }
}
