using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace Beporsoft.TabularSheets
{
    /// <summary>
    /// Represent the base class to build tabular sheets
    /// </summary>
    /// <typeparam name="T">The type which represent every row</typeparam>
    public abstract class TabularData<T> : IList<T>
    {
        internal List<TabularDataColumn<T>> ColumnsCollection = new();
        private List<T> _items = new();

        /// <summary></summary>
        public TabularData()
        {
        }

        #region Properties
        /// <summary>
        /// Gets the collection of items which will be displayed on rows.
        /// </summary>
        public IList<T> Items { get => _items; protected set => _items = value.ToList(); }

        /// <summary>
        /// Gets the collection of columns which will define which data is displayed on each row, each cell.
        /// </summary>
        public IEnumerable<TabularDataColumn<T>> Columns => ColumnsCollection.AsReadOnly();

        /// <inheritdoc cref="List{T}.Count"/>
        public int Count => _items.Count;

        /// <inheritdoc cref="ICollection{T}.IsReadOnly"/>
        bool ICollection<T>.IsReadOnly => false;


        /// <summary>
        /// Get or sets the element at the specified index.
        /// </summary>
        /// <param name="index">The zero-based index of the element to get or set.</param>
        /// <returns></returns>
        public T this[int index] { get => _items[index]; set => _items[index] = value; }
        #endregion

        #region Manipulate Rows

        #region Interface ICollection
        /// <inheritdoc cref="IList{T}.IndexOf(T)"/>
        public int IndexOf(T item) => _items.IndexOf(item);

        /// <inheritdoc cref="IList{T}.Insert(int, T)"/>
        public void Insert(int index, T item) => _items.Insert(index, item);

        /// <inheritdoc cref="IList{T}.RemoveAt(int)"/>
        public void RemoveAt(int index) => _items.RemoveAt(index);

        /// <inheritdoc cref="ICollection{T}.Add(T)"/>
        public void Add(T item) => _items.Add(item);

        /// <inheritdoc cref="ICollection{T}.Clear"/>
        public void Clear() => _items.Clear();

        /// <inheritdoc cref="ICollection{T}.Contains(T)"/>
        public bool Contains(T item) => _items.Contains(item);

        /// <inheritdoc cref="ICollection{T}.CopyTo(T[], int)"/>
        public void CopyTo(T[] array, int arrayIndex) => _items.CopyTo(array, arrayIndex);

        /// <inheritdoc cref="ICollection{T}.Remove(T)"/>
        public bool Remove(T item) => _items.Remove(item);

        /// <inheritdoc cref="IEnumerable{T}.GetEnumerator"/>
        public IEnumerator<T> GetEnumerator() => _items.GetEnumerator();

        /// <inheritdoc cref="IEnumerable.GetEnumerator"/>
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
            ColumnsCollection.Add(column);
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
            ColumnsCollection.Add(column);
            return column;
        }


        /// <summary>
        /// Remove the given column if exists and reorganize the order of the following columns
        /// </summary>
        /// <param name="column">The column to remove</param>
        /// <returns><see langword="true"/> if the column was removed. Otherwise, or if it doesn't exist on collection, <see langword="false"/></returns>
        public bool RemoveColumn(TabularDataColumn<T> column)
        {
            bool removed = ColumnsCollection.Remove(column);
            //if (removed)
            //{
            //    // the next column and above will reorganize the column order.
            //    int columnPosition = column.Order;
            //    TabularDataColumn<T>? nextColumn = Columns.SingleOrDefault(c => c.Order == columnPosition + 1);
            //    nextColumn?.SetPosition(columnPosition);
            //}
            return removed;
        }
        #endregion

        #region Helpers Columns
        /// <summary>
        /// Method invoked from columns which belongs to this table to reorganize the list of column order
        /// </summary>
        //internal void ReallocateColumn(TabularDataColumn<T> column, int newPosition)
        //{
        //    if (newPosition > Columns.Count)
        //        throw new ArgumentOutOfRangeException(nameof(newPosition), newPosition, "The value of position cannot be higher than the amount of columns");
        //    Columns.Remove(column);
        //    var cols = Columns.OrderBy(x => x.Order);
        //    int idx = 0;
        //    foreach (var col in cols)
        //    {
        //        if (idx >= newPosition)
        //            col.Order = idx + 1;
        //        else
        //            col.Order = idx;
        //        idx++;
        //        col.SetDefaultName();
        //    }
        //    column.Order = newPosition;
        //    column.SetDefaultName();
        //    Columns.Add(column);
        //}
        #endregion

        internal object Evaluate(int row, int col)
        {
            if (Count >= row)
                throw new ArgumentOutOfRangeException(nameof(row), row, $"The value of row is outside the bounds of the collection length");
            if (ColumnsCollection.Count >= col)
                throw new ArgumentOutOfRangeException(nameof(col), col, $"The value of col is outside the bounds of the columns length");

            T item = this[row];
            TabularDataColumn<T> func = ColumnsCollection[col];

            object result = func.Apply(item);
            return result;
        }

    }
}
