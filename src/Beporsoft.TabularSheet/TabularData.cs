using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheet
{
    public class TabularData<T> : List<T>
    {

        public TabularData()
        {
        }

        public ICollection<TabularDataColumn<T>> Columns => Cols;
        internal List<TabularDataColumn<T>> Cols { get; } = new List<TabularDataColumn<T>>();

        public TabularData<T> SetColumn(Func<T, object> predicate)
        {
            var column = new TabularDataColumn<T>(this, predicate);
            int order = Cols.Any() ? (Cols.Max(x => x.Order) + 1) : 0;
            Cols.Add(column);
            return this;
        }

        public object Evaluate(int row, int col)
        {
            if (this.Count >= row)
                throw new ArgumentOutOfRangeException(nameof(row), row, $"The value of row is outside the bounds of the collection length");
            if (Cols.Count >= col)
                throw new ArgumentOutOfRangeException(nameof(col), col, $"The value of col is outside the bounds of the columns length");

            T item = this[row];
            TabularDataColumn<T> func = Cols[col];

            object result = func.Apply(item);
            return result;
        }
    }
}
