using Beporsoft.TabularSheets.Style;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets
{
    /// <summary>
    /// Represent a column inside a <see cref="TabularData{T}"/>
    /// </summary>
    /// <typeparam name="T">The type of the instances which will be represented for each row</typeparam>
    [DebuggerDisplay("{Title} | {Order} | {ColumnData}")]
    public class TabularDataColumn<T>
    {
        private const string _defaultColumnName = "Column"; // Default name for unnamed columns, adding order number.
        private static readonly Regex _regexDefaultColumnName = new(_defaultColumnName + @"\d{0,}"); // Regex to find when a column is named according to its default name.

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

        /// <summary>
        /// The function which will be evaluated to fill the respective column for each item.
        /// </summary>
        public Func<T, object> ColumnData { get; init; }

        /// <summary>
        /// The title of the column. It must be set throught constructor or <see cref="SetTitle(string)"/>
        /// </summary>
        public string Title { get; private set; } = null!;

        /// <summary>
        /// The position where the column will be displayed.
        /// </summary>
        public int Order { get; private set; } = 0;
        public ColumnOptions ColumnOptions { get; set; } = new();
        public TabularData<T> Owner { get; }


        public TabularDataColumn<T> SetTitle(string title)
        {
            Title = title;
            return this;
        }

        public object Apply(T value)
        {
            return ColumnData.Invoke(value);
        }

        /// <summary>
        /// Reassign the current order position of the column inside the parennt table
        /// </summary>
        public void SetPosition(int position)
        {
            if (position > Owner.Columns.Count)
                throw new ArgumentOutOfRangeException(nameof(position), position, "The value of position cannot be higher than the amount of columns");
            Owner.Columns.Remove(this);
            var cols = Owner.Columns.OrderBy(x => x.Order);
            int idx = 0;
            foreach (var col in cols)
            {
                if (idx >= position)
                    col.Order = idx + 1;
                else
                    col.Order = idx;
                idx++;
                col.SetDefaultName();
            }
            Order = position;
            SetDefaultName();
            Owner.Columns.Add(this);
        }

        /// <summary>
        /// When no title is provided, set the column title to {_defaultColumnName}{Order}.
        /// </summary>
        internal void SetDefaultName()
        {
            Match match = _regexDefaultColumnName.Match(Title ?? string.Empty); // Avoid Exception when title null
            if (string.IsNullOrWhiteSpace(Title) || match.Success)
            {
                Title = $"{_defaultColumnName}{Order}";
            }
        }

    }
}
