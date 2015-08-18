using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel.Helper
{
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class ExcelColumnAttribute : Attribute
    {
        private readonly string _columnName;

        public ExcelColumnAttribute(string columnName)
        {
            _columnName = columnName;
        }

        public string ColumnName
        {
            get { return _columnName; }
        }
    }
}
