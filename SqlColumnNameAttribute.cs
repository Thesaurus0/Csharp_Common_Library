using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommonLib
{
    public class SqlColumnNameAttribute:  System.Attribute
    {
        public SqlColumnNameAttribute(string sqlColumnName)
        {
            this._sqlColName = sqlColumnName;
        }
        private string _sqlColName;

        public string SqlColumnName
        {
            get { return _sqlColName; }
        }

        public string ColumnFormat { get; set; }
    }
}
