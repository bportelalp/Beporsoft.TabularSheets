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
        private List<T> _items = new();

        #region Properties
        /// <summary>
        /// Gets the collection of items which will be displayed on rows.
        /// </summary>
        public IList<T> Items { get => _items; protected set => _items = value.ToList(); }

        /// <summary>
        /// Gets the collection of columns that will define which data is displayed on each row, each cell.
        /// </summary>
        public IEnumerable<TabularDataColumn<T>> Columns => ColumnsCollection.AsReadOnly();

        /// <summary>
        /// Gets the number of rows.
        /// </summary>
        public int Count => _items.Count;

        /// <summary>
        /// Gets the number of columns
        /// </summary>
        public int ColumnCount => ColumnsCollection.Count;

        /// <inheritdoc cref="ICollection{T}.IsReadOnly"/>
        bool ICollection<T>.IsReadOnly => false;

        /// <summary>
        /// Collection to be used from the children. Internal to the assembly
        /// </summary>
        internal List<TabularDataColumn<T>> ColumnsCollection { get; } = new();

        /// <summary>
        /// Get or sets the element at the specified index.
        /// </summary>
        /// <param name="index">The zero-based index of the element to get or set.</param>
        /// <returns></returns>
        public T this[int index] { get => _items[index]; set => _items[index] = value; }
        #endregion

        #region Manipulate Rows

        #region Interface IList

        /// <inheritdoc cref="ICollection{T}.Add(T)"/>
        public void Add(T item) => _items.Add(item);

        /// <inheritdoc cref="ICollection{T}.Clear"/>
        public void Clear() => _items.Clear();

        /// <inheritdoc cref="ICollection{T}.Contains(T)"/>
        public bool Contains(T item) => _items.Contains(item);

        /// <inheritdoc cref="ICollection{T}.CopyTo(T[], int)"/>
        public void CopyTo(T[] array, int arrayIndex) => _items.CopyTo(array, arrayIndex);

        /// <inheritdoc cref="IList{T}.IndexOf(T)"/>
        public int IndexOf(T item) => _items.IndexOf(item);

        /// <inheritdoc cref="IList{T}.Insert(int, T)"/>
        public void Insert(int index, T item) => _items.Insert(index, item);

        /// <inheritdoc cref="ICollection{T}.Remove(T)"/>
        public bool Remove(T item) => _items.Remove(item);

        /// <inheritdoc cref="IList{T}.RemoveAt(int)"/>
        public void RemoveAt(int index) => _items.RemoveAt(index);

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
        /// Adds a new column with the information that will be displayed on it.
        /// </summary>
        /// <param name="predicate">A expression which will be evaluated to populate the cell</param>
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
        /// <param name="predicate">A expression which will be evaluated to populate the cell</param>
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
            return removed;
        }
        #endregion

        /// <summary>
        /// Evaluate the value of a cell inside the body table, that is, excluding the header
        /// </summary>
        /// <param name="row">zero based index of the cell</param>
        /// <param name="col"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        protected object EvaluateCell(int row, int col)
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
