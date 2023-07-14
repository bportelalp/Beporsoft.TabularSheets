using Beporsoft.TabularSheets.Builders.SheetBuilders;
using Beporsoft.TabularSheets.Builders.StyleBuilders;
using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace Beporsoft.TabularSheets.Builders.Interfaces
{
    /// <summary>
    /// Provides an abstraction for making the adequate OpenXml<see cref="SheetData"/> which represent the current instance. <br/>
    /// The <see cref="SheetData"/> node represent all the data contained in one sheet of the spreadsheets. This abstraction will allow in 
    /// future implementations that one spreadsheet will contain more than one <see cref="ISheet"/>, even with different type.
    /// </summary>
    internal interface ISheet
    {
        /// <summary>
        /// Title of the sheet
        /// </summary>
        internal string Title { get; }

        /// <summary>
        /// The system type which contains this <see cref="ISheet"/>
        /// </summary>
        internal Type ItemType { get; }

        /// <summary>
        /// Build the object <see cref="WorksheetBundle"/> which contains all the elements to build a <see cref="Worksheet"/>
        /// and fill <paramref name="stylesheetBuilder"/> and <paramref name="sharedStringBuilder"/> with the styles and 
        /// shared strings discovered during the creation, respectively
        /// </summary>
        /// <param name="stylesheetBuilder">A reference to the object which handles the compilation of stylesheet</param>
        /// <param name="sharedStringBuilder">A reference to the object which handles the compilation of shared strings</param>
        /// <returns></returns>
        internal WorksheetBundle BuildSheetContext(StylesheetBuilder stylesheetBuilder, SharedStringBuilder sharedStringBuilder);


    }
}
