using Beporsoft.TabularSheets.CellStyling;
using DocumentFormat.OpenXml.Office2016.Drawing.Command;
using System;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;

namespace Beporsoft.TabularSheets
{
    /// <summary>
    /// Represent a column inside a <see cref="TabularData{T}"/>
    /// </summary>
    /// <typeparam name="T">The type of the instances which will be represented for each row</typeparam>
    [DebuggerDisplay("{Title} | {Order} | {ColumnData}")]
    public class TabularDataColumn<T>
    {
        private const string _defaultColumnName = "Col"; // Default name for unnamed columns, of pattern {typeof(T).Name}{_defautl}


        private string? _title;

        #region Constructors
        internal TabularDataColumn(TabularData<T> parentTabularData, Func<T, object> columnData)
        {
            Owner = parentTabularData;
            ColumnContent = columnData;
            //Order = Owner.Columns.Any() ? Owner.Columns.Max(x => x.Order) + 1 : 0;
            //SetDefaultName();
        }

        internal TabularDataColumn(string title, TabularData<T> parentTabularData, Func<T, object> columnData) : this(parentTabularData, columnData)
        {
            _title = title;
        }
        #endregion

        #region Properties
        /// <summary>
        /// Gets or sets the function which will be evaluated to fill the respective column for each item.
        /// </summary>
        public Func<T, object> ColumnContent { get; set; }

        /// <summary>
        /// Gets title of the column. If no title are provided before, the default one is displayed.<br/>
        /// To set it, use the method <see cref="SetTitle(string)"/>
        /// </summary>
        public string Title => GetTitle();

        /// <summary>
        /// Gets the column index of the column on the parent table
        /// </summary>
        public int ColumnIndex => Owner.ColumnsCollection.IndexOf(this);

        /// <summary>
        /// Gets the style to apply to all the column.
        /// </summary>
        public Style Style { get; private set; } = Style.Default;

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
            _title = title;
            return this;
        }


        /// <summary>
        /// Set the style of all the column. Header is excluded
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
            // take the current style, or create a new one to edit fields
            Style editionStyle = Style;
            if (Style.Equals(Style.Default))
                editionStyle = new();

            styleActionEdition.Invoke(editionStyle);
            // If it is equals default, set the default again
            if (editionStyle.Equals(Style.Default))
                Style = Style.Default;

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
                throw new ArgumentOutOfRangeException(nameof(index), $"{nameof(index)} is less than 0 or is greater than the current amount of items");
            Owner.ColumnsCollection.Remove(this);
            Owner.ColumnsCollection.Insert(index, this);
            return this;
        }
        #endregion

        #region Internal methods called from Sheet builders
        /// <summary>
        /// Retrieve the current value of the cell for this column applied to <paramref name="value"/>
        /// </summary>
        internal object Apply(T value)
        {
            return ColumnContent.Invoke(value);
        }

        /// <summary>
        /// Gets the defined title or build a default one if it is empry
        /// </summary>
        internal string GetTitle()
        {
            if (string.IsNullOrWhiteSpace(_title))
            {
                string format = $"D{Owner.Columns.Count().ToString().Count()}"; // Get the significant digits of the columns count, to give format
                return $"{typeof(T).Name}{_defaultColumnName}{ColumnIndex.ToString(format)}";
            }
            else
            {
                return _title!;
            }
        }
        #endregion

    }
}
