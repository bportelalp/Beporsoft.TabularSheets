using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheet
{
    /// <summary>
    /// Represent a column inside a <see cref="TabularData{T}"/>
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class TabularDataColumn<T>
    {

        internal TabularDataColumn(TabularData<T> parentTabularData, Func<T, object> columnData)
        {
            TabularDataOwner = parentTabularData;
            ColumnData = columnData;
        }

        public Func<T, object> ColumnData { get; set; }
        public string Title { get; set; } = null!;
        public int Order { get; private set; }
        public TabularData<T> TabularDataOwner { get; }


        public object Apply(T value)
        {
            return ColumnData.Invoke(value);
        }

        public TabularDataColumn<T> ChangeOrder(int order)
        {
            Order = order;
            return this;
        }


    }
}
