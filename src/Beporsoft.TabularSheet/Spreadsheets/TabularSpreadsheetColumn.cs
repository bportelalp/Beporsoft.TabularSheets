using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheet.Spreadsheets
{
    public class TabularSpreadsheetColumn<T> : TabularDataColumn<T>
    {
        internal TabularSpreadsheetColumn(TabularData<T> parentTabularData, Func<T, object> columnData) : base(parentTabularData, columnData)
        {
        }

        public int ColumnWidth { get; set; }
    }
}
