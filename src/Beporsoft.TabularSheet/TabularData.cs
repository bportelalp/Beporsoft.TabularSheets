using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheet
{
    public class TabularData<T> : IList<T>
    {
        private readonly List<T> _items = new();
        internal readonly List<TabularDataColumn<T>> _columns = new();

        public TabularData()
        {
        }

        public ICollection<TabularDataColumn<T>> Columns => _columns;
        public int Count => _items.Count;
        public bool IsReadOnly => false;

        public T this[int index] { get => _items[index]; set => _items[index] = value; }


        #region Manipulate Rows

        #region Interface IList
        public int IndexOf(T item) => _items.IndexOf(item);

        public void Insert(int index, T item) => _items.Insert(index, item);

        public void RemoveAt(int index) => _items.RemoveAt(index);

        public void Add(T item) => _items.Add(item);

        public void Clear() => _items.Clear();

        public bool Contains(T item) => _items.Contains(item);

        public void CopyTo(T[] array, int arrayIndex) => _items.CopyTo(array, arrayIndex);

        public bool Remove(T item) => _items.Remove(item);

        public IEnumerator<T> GetEnumerator() => _items.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
        #endregion

        /// <inheritdoc cref="List{T}.AddRange(IEnumerable{T})"/>
        public void AddRange(IEnumerable<T> items) => _items.AddRange(items);
        #endregion

        #region Manipulate Columns
        public TabularDataColumn<T> SetColumn(Func<T, object> predicate)
        {
            var column = new TabularDataColumn<T>(this, predicate);
            int order = _columns.Any() ? (_columns.Max(x => x.Order) + 1) : 0;
            _columns.Add(column);
            return column;
        }

        public bool RemoveColumn(TabularDataColumn<T> column) => _columns.Remove(column);
        #endregion

        public object Evaluate(int row, int col)
        {
            if (this.Count >= row)
                throw new ArgumentOutOfRangeException(nameof(row), row, $"The value of row is outside the bounds of the collection length");
            if (_columns.Count >= col)
                throw new ArgumentOutOfRangeException(nameof(col), col, $"The value of col is outside the bounds of the columns length");

            T item = this[row];
            TabularDataColumn<T> func = _columns[col];

            object result = func.Apply(item);
            return result;
        }

    }
}
