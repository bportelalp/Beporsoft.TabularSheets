using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace Beporsoft.TabularSheets
{
    public abstract class TabularData<T> : IList<T>
    {
        internal readonly List<TabularDataColumn<T>> _columns = new();
        private List<T> _items = new();

        public TabularData()
        {
        }

        /// <summary>
        /// Gets the collection of items which will be displayed on rows
        /// </summary>
        public ICollection<T> Items { get => _items; protected set => _items = value.ToList(); }

        /// <summary>
        /// Gets the collection of columns which will define which data is displayed on each row, each cell.
        /// </summary>
        public virtual ICollection<TabularDataColumn<T>> Columns => _columns;
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
        /// <summary>
        /// Adds a new column which the information that will be displayed on a column.
        /// </summary>
        /// <param name="predicate">A expression which will be evaluated for fill the cell</param>
        /// <returns>The column created, so additional calls can be chained</returns>
        public virtual TabularDataColumn<T> AddColumn(Func<T, object> predicate)
        {
            var column = new TabularDataColumn<T>(this, predicate);
            _columns.Add(column);
            return column;
        }

        /// <summary>
        /// <inheritdoc cref="AddColumn(Func{T, object})"/>
        /// </summary>
        /// <param name="title">The title of the column</param>
        /// <param name="predicate">A expression which will be evaluated for fill the cell</param>
        /// <returns>The column created, so additional calls can be chained</returns>
        public virtual TabularDataColumn<T> AddColumn(string title, Func<T, object> predicate)
        {
            var column = new TabularDataColumn<T>(title, this, predicate);
            _columns.Add(column);
            return column;
        }

        /// <summary>
        /// Remove the given column if exists and reorganize the order of the following columns
        /// </summary>
        /// <param name="column">The column to remove</param>
        /// <returns><see langword="true"/> if the column was removed. Otherwise, or if it doesn't exist on collection, <see langword="false"/></returns>
        public bool RemoveColumn(TabularDataColumn<T> column)
        {
            bool removed = Columns.Remove(column);
            if (removed)
            {
                // the next column and avobe will reorganize the column order.
                int columnPosition = column.Order;
                TabularDataColumn<T>? nextColumn = Columns.SingleOrDefault(c => c.Order == columnPosition + 1);
                nextColumn?.SetPosition(columnPosition);
            }
            return removed;
        }
        #endregion


        /// <summary>
        /// Method invoked from columns which belongs to this table to reorganize the list of column order
        /// </summary>
        internal void ReallocateColumn(TabularDataColumn<T> column, int newPosition)
        {
            if (newPosition > Columns.Count)
                throw new ArgumentOutOfRangeException(nameof(newPosition), newPosition, "The value of position cannot be higher than the amount of columns");
            Columns.Remove(column);
            var cols = Columns.OrderBy(x => x.Order);
            int idx = 0;
            foreach (var col in cols)
            {
                if (idx >= newPosition)
                    col.Order = idx + 1;
                else
                    col.Order = idx;
                idx++;
                col.SetDefaultName();
            }
            column.Order = newPosition;
            column.SetDefaultName();
            Columns.Add(column);
        }


        public object Evaluate(int row, int col)
        {
            if (Count >= row)
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
