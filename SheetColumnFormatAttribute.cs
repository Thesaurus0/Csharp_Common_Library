using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommonLib {
    [AttributeUsage(AttributeTargets.Property, AllowMultiple =false, Inherited =false)]
    public class SheetColumnFormatAttribute : System.Attribute {
        private string _format;
        public string ColumnFormat {
            get { return _format; }
        }

        public SheetColumnFormatAttribute(string format) {
            this._format = format;
        }
    }
}
