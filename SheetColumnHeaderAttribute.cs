using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommonLib
{
    public class SheetColumnHeaderAttribute : System.Attribute
    {
        private string _header;

        public string Header {
            get { return _header; }
        }

        public SheetColumnHeaderAttribute(string header) {
            this._header = header;
        }
    }
}
