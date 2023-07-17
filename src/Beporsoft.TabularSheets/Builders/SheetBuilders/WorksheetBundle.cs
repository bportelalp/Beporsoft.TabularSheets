using DocumentFormat.OpenXml.Spreadsheet;

namespace Beporsoft.TabularSheets.Builders.SheetBuilders
{
    /// <summary>
    /// A bundle which contains all the elements to build a <see cref="Worksheet"/>
    /// </summary>
    internal class WorksheetBundle
    {
        /// <summary>
        /// Container of the data by rows
        /// </summary>
        public SheetData Data { get; set; } = null!;

        /// <summary>
        /// Container of the style of each column. 
        /// </summary>
        public Columns? Columns { get; set; }

        /// <summary>
        /// Container of general properties about the formatting of the sheet
        /// </summary>
        public SheetFormatProperties FormatProperties { get; set; } = null!;
    }
}
