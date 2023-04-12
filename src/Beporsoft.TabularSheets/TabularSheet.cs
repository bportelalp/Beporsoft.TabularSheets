using Beporsoft.TabularSheets.Builders;
using Beporsoft.TabularSheets.Builders.Interfaces;
using Beporsoft.TabularSheets.Builders.SheetBuilders;
using Beporsoft.TabularSheets.Builders.StyleBuilders;
using Beporsoft.TabularSheets.CellStyling;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Beporsoft.TabularSheets
{
    public sealed class TabularSheet<T> : TabularData<T>, ISheet
    {
        public TabularSheet()
        {
        }

        public TabularSheet(IEnumerable<T> items)
        {
            Items = items.ToList();
        }

        #region Properties
        /// <summary>
        /// The title of the current sheet
        /// </summary>
        public string Title { get; set; } = "Sheet";

        /// <summary>
        /// Defines the style of heading cells of <see cref="TabularSheet{T}"/>.<br/> 
        /// Initially all styles are the default ones from <see cref="DefaultStyle"/>
        /// </summary>
        public Style HeaderStyle { get; private set; } = new();

        /// <summary>
        /// Defines the style of data cells from the current <see cref="TabularSheet{T}"/>. Applies to
        /// all cells unless a more specific style is applied.
        /// </summary>
        public DefaultStyle DefaultStyle { get; private set; } = new();

        /// <summary>
        /// Enable the automatic column filter
        /// </summary>
        public bool AutoFilter { get; set; } = false;
        #endregion

        #region Configure Table
        public void SetSheetTitle(string title) => Title = title;
        #endregion

        #region Create
        public void Create(string path)
        {
            SpreadsheetBuilder builder = new();
            builder.Create(path, this);
        }

        public MemoryStream Create()
        {
            SpreadsheetBuilder builder = new();
            MemoryStream stream = builder.Create(this);
            return stream;
        }
        #endregion

        #region ISheet
        Type ISheet.ItemType => typeof(T);

        SheetData ISheet.BuildData(ref StylesheetBuilder stylesheetBuilder, ref SharedStringBuilder sharedStringBuilder)
        {
            SheetBuilder<T> builder = new(this, stylesheetBuilder, sharedStringBuilder);
            return builder.BuildSheetData();
        }

        Columns ISheet.BuildColumns()
        {
            return SheetBuilder<T>.BuildColumns(this);
        }
        #endregion
    }
}
