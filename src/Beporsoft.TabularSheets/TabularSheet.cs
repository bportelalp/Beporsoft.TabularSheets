using Beporsoft.TabularSheets.Builders;
using Beporsoft.TabularSheets.Builders.Interfaces;
using Beporsoft.TabularSheets.Builders.SheetBuilders;
using Beporsoft.TabularSheets.Builders.StyleBuilders;
using Beporsoft.TabularSheets.CellStyling;
using Beporsoft.TabularSheets.Options;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Beporsoft.TabularSheets
{
    /// <summary>
    /// Represent a spreadsheet that can be handled by the OpenXml Specification
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public sealed class TabularSheet<T> : TabularData<T>, ISheet
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="TabularSheet{T}"/> class that is empty
        /// </summary>
        public TabularSheet()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="TabularSheet{T}"/> class that contains elements copied from the specific collection.
        /// </summary>
        /// <param name="items">The collection whose elements are copied to the new tabular sheet</param>
        public TabularSheet(IEnumerable<T> items)
        {
            foreach (var item in items)
                Items.Add(item);
        }

        #region Properties
        /// <summary>
        /// Gets or sets the title of the current sheet
        /// </summary>
        public string Title { get; set; } = "Sheet";

        /// <summary>
        /// Gets the style of heading cells of <see cref="TabularSheet{T}"/>.
        /// </summary>
        public Style HeaderStyle { get; private set; } = new();

        /// <summary>
        /// Gets the style of data cells from the current <see cref="TabularSheet{T}"/>.
        /// </summary>
        public Style BodyStyle { get; private set; } = new();

        /// <summary>
        /// Gets the common options to configure the spreadsheet creation
        /// </summary>
        public SheetOptions Options { get; private set; } = new();
        #endregion

        #region Configure Table
        /// <summary>
        /// Set the name of the sheet
        /// </summary>
        /// <param name="title"></param>
        public void SetSheetTitle(string title) => Title = title;
        #endregion

        #region Create
        /// <summary>
        /// Create a spreadsheet document
        /// </summary>
        /// <param name="path">Path to store the document</param>
        public void Create(string path)
        {
            SpreadsheetBuilder builder = new();
            builder.Create(path, this);
        }

        /// <summary>
        /// Create a spreadsheet document
        /// </summary>
        public MemoryStream Create()
        {
            SpreadsheetBuilder builder = new();
            MemoryStream stream = builder.Create(this);
            return stream;
        }
        #endregion

        #region ISheet
        Type ISheet.ItemType => typeof(T);

        WorksheetBundle ISheet.BuildSheetContext(StylesheetBuilder stylesheetBuilder, SharedStringBuilder sharedStringBuilder)
        {
            SheetBuilder<T> builder = new(this, stylesheetBuilder, sharedStringBuilder);
            SheetData data = builder.BuildSheetData();
            Columns? cols = builder.BuildColumns();
            SheetFormatProperties formatProps = builder.BuildFormatProperties();

            return new WorksheetBundle()
            {
                Data =data,
                Columns =cols,
                FormatProperties = formatProps,
            };
        }
        #endregion
    }
}
