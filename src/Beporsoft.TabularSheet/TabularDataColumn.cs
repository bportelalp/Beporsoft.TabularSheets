using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheet
{
    /// <summary>
    /// Represent a column inside a <see cref="TabularData{T}"/>
    /// </summary>
    /// <typeparam name="T"></typeparam>
    [DebuggerDisplay("{Title} | {Order} | {ColumnData}")]
    public class TabularDataColumn<T>
    {
        private const string _defaultColumnName = "Column";

        internal TabularDataColumn(TabularData<T> parentTabularData, Func<T, object> columnData)
        {
            Owner = parentTabularData;
            ColumnData = columnData;
            Order = Owner.Columns.Any() ? (Owner.Columns.Max(x => x.Order) + 1) : 1;
            Title = $"{_defaultColumnName}{Order}";
        }

        internal TabularDataColumn(string title, TabularData<T> parentTabularData, Func<T, object> columnData) : this(parentTabularData, columnData)
        {
            Title = title;
        }

        public Func<T, object> ColumnData { get; set; }
        public string Title { get; private set; } = null!;
        public int Order { get; private set; } = 1;
        public TabularData<T> Owner { get; }


        public void SetTitle(string title)
        {
            Title = title;
        }

        public object Apply(T value)
        {
            return ColumnData.Invoke(value);
        }

    }
}
