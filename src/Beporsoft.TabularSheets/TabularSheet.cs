using Beporsoft.TabularSheets.Builders;
using Beporsoft.TabularSheets.Builders.Interfaces;
using Beporsoft.TabularSheets.Builders.SheetBuilders;
using Beporsoft.TabularSheets.Builders.StyleBuilders;
using Beporsoft.TabularSheets.Style;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Beporsoft.TabularSheets
{
    public class TabularSheet<T> : TabularData<T>, ISheet
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
        public HeaderStyle HeaderStyle { get; private set; } = new();
        public DefaultStyle Options { get; private set; } = new();

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
        #endregion
    }
}
