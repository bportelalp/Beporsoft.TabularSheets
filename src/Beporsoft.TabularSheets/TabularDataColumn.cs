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
        private static readonly Regex _regexDefaultColumnName = new(nameof(T) + _defaultColumnName + @"\d{0,}"); // Regex to find when a column is named according to its default name.

        internal TabularDataColumn(TabularData<T> parentTabularData, Func<T, object> columnData)
        {
            Owner = parentTabularData;
            ColumnData = columnData;
            Order = Owner.Columns.Any() ? Owner.Columns.Max(x => x.Order) + 1 : 0;
            SetDefaultName();
        }

        internal TabularDataColumn(string title, TabularData<T> parentTabularData, Func<T, object> columnData) : this(parentTabularData, columnData)
        {
            Title = title;
        }

        #region Properties
        /// <summary>
        /// The function which will be evaluated to fill the respective column for each item.
        /// </summary>
        public Func<T, object> ColumnData { get; set; }

        /// <summary>
        /// The title of the column. It must be set throught constructor or <see cref="SetTitle(string)"/>
        /// </summary>
        public string Title { get; private set; } = null!;

        /// <summary>
        /// The position where the column will be displayed.
        /// </summary>
        public int Order { get; internal set; } = 0;

        /// <summary>
        /// The style to apply to all the column excluding the header
        /// </summary>
        public Style Style { get; set; } = Style.Default;

        /// <summary>
        /// The <see cref="TabularData{T}"/> to whom belongs this <see cref="TabularDataColumn{T}"/>
        /// </summary>
        public TabularData<T> Owner { get; }
        #endregion

        #region Edition
        /// <summary>
        /// Sets the title of the column
        /// </summary>
        /// <param name="title"></param>
        /// <returns>The same column instance, so additional calls can be chained</returns>
        public TabularDataColumn<T> SetTitle(string title)
        {
            Title = title;
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

        /// <summary>
        /// Reassign the current order position of the column inside the parent table and reorganice the previous items
        /// </summary>
        /// <returns>The same column instance, so additional calls can be chained</returns>
        public TabularDataColumn<T> SetPosition(int position)
        {
            Owner.ReallocateColumn(this, position);
            return this;
        }
        #endregion

        #region Internal methods called from Sheet builders
        /// <summary>
        /// Retrieve the current value of the cell for this column applied to <paramref name="value"/>
        /// </summary>
        internal object Apply(T value)
        {
            return ColumnData.Invoke(value);
        }

        /// <summary>
        /// When no title is provided, set the column title to {_defaultColumnName}{Order}.
        /// </summary>
        internal void SetDefaultName()
        {
            Match match = _regexDefaultColumnName.Match(Title ?? string.Empty);
            if (string.IsNullOrWhiteSpace(Title) || match.Success)
            {
                Title = $"{nameof(T)}{_defaultColumnName}{Order}";
            }
        }
        #endregion

    }
}
