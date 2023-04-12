using Beporsoft.TabularSheets.Builders.SheetBuilders;
using Beporsoft.TabularSheets.Builders.StyleBuilders;
using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace Beporsoft.TabularSheets.Builders.Interfaces
{
    /// <summary>
    /// Provides an abstraction of making the adequate <see cref="SheetData"/> which represent the current instance
    /// </summary>
    internal interface ISheet
    {
        internal string Title { get; }

        internal Type ItemType { get; }

        /// <summary>
        /// Build the object <see cref="SheetData"/> which represent the data.
        /// </summary>
        /// <param name="stylesheetBuilder">A reference to the object which handles the compilation of stylesheet</param>
        /// <param name="sharedStringBuilder">A reference to the object which handles the compilation of shared strings</param>
        /// <returns></returns>
        internal SheetData BuildData(ref StylesheetBuilder stylesheetBuilder, ref SharedStringBuilder sharedStringBuilder);


        internal Columns BuildColumns();
    }
}
