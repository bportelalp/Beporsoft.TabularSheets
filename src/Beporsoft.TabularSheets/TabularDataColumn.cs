using Beporsoft.TabularSheets.CellStyling;
using Beporsoft.TabularSheets.Options;
using System;
using System.Diagnostics;
using System.Linq;

namespace Beporsoft.TabularSheets
{
    /// <summary>
    /// Represent a column inside a <see cref="TabularData{T}"/>
    /// </summary>
    /// <typeparam name="T">The type of the instances which will be represented for each row</typeparam>
    [DebuggerDisplay("{Title} | Col={Index} | {CellContent}")]
    public class TabularDataColumn<T>
    {
        private const string _defaultColumnName = "Col"; // Default name for unnamed columns, of pattern {typeof(T).Name}{_default}
        private string? _columnTitle;

        #region Constructors
        internal TabularDataColumn(TabularData<T> parentTabularData, Func<T, object> columnData)
        {
            Owner = parentTabularData;
            CellContent = columnData;
        }

        internal TabularDataColumn(string title, TabularData<T> parentTabularData, Func<T, object> columnData) : this(parentTabularData, columnData)
        {
            _columnTitle = title;
        }
        #endregion

        #region Properties
        /// <summary>
        /// Gets the function which will be evaluated to fill the respective column for each item.
        /// </summary>
        public Func<T, object> CellContent { get; private set; }

        /// <summary>
        /// Gets title of the column. If no title are provided before, the default one is displayed.<br/>
        /// To set it, use the method <see cref="SetTitle(string)"/>
        /// </summary>
        public string Title => GetTitle();

        /// <summary>
        /// Gets the column index of the column on the parent table
        /// </summary>
        public int Index => Owner.ColumnsCollection.IndexOf(this);

        /// <summary>
        /// Gets the style to apply to all the column.
        /// </summary>
        public Style Style { get; private set; } = new();

        internal ColumnOptions Options { get; private set; } = new();

        /// <summary>
        /// Gets the <see cref="TabularData{T}"/> to which belongs this <see cref="TabularDataColumn{T}"/>
        /// </summary>
        public TabularData<T> Owner { get; }

        #endregion

        #region Edition
        /// <summary>
        /// Sets the title of the column
        /// </summary>
        /// <param name="title">The title to assign to the column, or <see langword="null"/> or <see cref="string.Empty"/> for the default one</param>
        /// <returns>The same column instance, so additional calls can be chained</returns>
        public TabularDataColumn<T> SetTitle(string? title)
        {
            _columnTitle = title;
            return this;
        }


        /// <summary>
        /// Set the style for all the column. Header is excluded
        /// </summary>
        /// <param name="style"></param>
        /// <returns>The same column instance, so additional calls can be chained</returns>
        public TabularDataColumn<T> SetStyle(Style style)
        {
            Style = style;
            return this;
        }

        /// <inheritdoc cref="SetStyle(Style)"/>
        /// <param name="styleActionEdition">An action to use the given object and edit the fields</param>
        /// <returns></returns>
        public TabularDataColumn<T> SetStyle(Action<Style> styleActionEdition)
        {
            styleActionEdition.Invoke(Style);
            return this;
        }

        /// <summary>
        /// Reassign the current order position of the column inside the parent table and reorganice the previous items
        /// </summary>
        /// <param name="index"></param>
        /// <returns>The same column instance, so additional calls can be chained</returns>
        public TabularDataColumn<T> SetIndex(int index)
        {
            if (index < 0 || index >= Owner.ColumnsCollection.Count)
                throw new ArgumentOutOfRangeException(nameof(index), $"{nameof(index)} is less than 0 or is greater than the current amount of columns");
            Owner.ColumnsCollection.Remove(this);
            Owner.ColumnsCollection.Insert(index, this);
            return this;
        }
        #endregion

        #region Internal methods called from Sheet builders
        /// <summary>
        /// Retrieve the current value of the cell for this column applied to <paramref name="value"/>
        /// </summary>
        internal object Apply(T value) => CellContent.Invoke(value);

        /// <summary>
        /// Gets the defined title or build a default one if it is empry
        /// </summary>
        private string GetTitle()
        {
            if (string.IsNullOrWhiteSpace(_columnTitle))
            {
                string format = $"D{Owner.Columns.Count().ToString().Length}"; // Get the significant digits of the columns count, to give format
                return $"{typeof(T).Name}{_defaultColumnName}{Index.ToString(format)}";
            }
            else
            {
                return _columnTitle!;
            }
        }
        #endregion

    }
}
