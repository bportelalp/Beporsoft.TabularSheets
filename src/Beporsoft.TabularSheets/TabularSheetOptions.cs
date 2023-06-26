using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.CellStyling
{
    /// <summary>
    /// A collection of common options during build of Spreadsheet
    /// </summary>
    public class TabularSheetOptions
    {
        /// <summary>
        /// The default Date time format applied to cells which contains a <see cref="DateTime"/> value.<br/>
        /// By default the pattern is [ShortDatePattern LongTimePattern] regarding the current culture
        /// <br/><br/>
        /// NB: OpenXml commonly stores the datetime as double and use the <see cref="DocumentFormat.OpenXml.Spreadsheet.CellValues.Number"/> 
        /// as the default datatype. This conversion can be achieved using <see cref="DateTime.ToOADate"/>. The
        /// representation over the cell is commonly exposed using minimal stylesheet to show the common format. This field
        /// is required to apply as the default <see cref="Style.NumberingPattern"/> to every cell which contains a <see cref="DateTime"/> type, if any more
        /// specific style is applied on the styling properties.
        /// </summary>
        public string DateTimeFormat { get; set; } = $"{CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern} {CultureInfo.CurrentCulture.DateTimeFormat.LongTimePattern}";

        /// <summary>
        /// When <see langword="true"/>, <see cref="TabularSheet{T}.HeaderStyle"/> null properties will be overrided from
        /// the equivalent <see cref="TabularSheet{T}.BodyStyle"/> properties
        /// </summary>
        public bool InheritHeaderStyleFromBody { get; set; } = false;
    }
}
