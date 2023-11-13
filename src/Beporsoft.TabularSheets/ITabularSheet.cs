using Beporsoft.TabularSheets.Builders.SheetBuilders;
using Beporsoft.TabularSheets.Builders.StyleBuilders;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;

namespace Beporsoft.TabularSheets
{
    /// <summary>
    /// Represent a simple, non generic <see cref="TabularSheet{T}"/>.
    /// </summary>
    public interface ITabularSheet
    {
        /// <summary>
        /// Gets the title of the <see cref="ITabularSheet"/>
        /// </summary>
        public string Title { get; }

        /// <summary>
        /// Gets the type of data which will populate the <see cref="ITabularSheet"/> rows.
        /// </summary>
        public Type ItemType { get; }


        /// <summary>
        /// Build the object <see cref="Worksheet"/> used on the spreadsheet document
        /// and fill <paramref name="stylesheetBuilder"/> and <paramref name="sharedStringBuilder"/> with the styles and 
        /// shared strings discovered during the creation, respectively
        /// </summary>
        /// <param name="stylesheetBuilder">A reference to the object which handles the compilation of stylesheet</param>
        /// <param name="sharedStringBuilder">A reference to the object which handles the compilation of shared strings</param>
        /// <returns></returns>
        internal Worksheet BuildWorksheet(StylesheetBuilder stylesheetBuilder, SharedStringBuilder sharedStringBuilder);


    }
}
