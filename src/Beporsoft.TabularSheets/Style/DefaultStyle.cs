using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Style
{
    /// <summary>
    /// Defines the default styles for every cell of the representation of <see cref="TabularSheet{T}"/> inside
    /// a Spreadsheet Document.<br/> Initially all the styles are the default ones
    /// </summary>
    public class DefaultStyle
    {
        /// <summary>
        /// The default <see cref="FontStyle"/> applied to every cell of <see cref="TabularSheet{T}"/> if there is no more specific one
        /// </summary>
        public FontStyle DefaultFont { get; set; } = new FontStyle();

        /// <summary>
        /// The default <see cref="FillStyle"/> applied to every cell of <see cref="TabularSheet{T}"/> if there is no more specific one
        /// </summary>
        public FillStyle DefaultFill { get; set; } = new FillStyle();

        /// <summary>
        /// The default <see cref="BorderStyle"/> applied to every cell of <see cref="TabularSheet{T}"/> if there is no more specific one
        /// </summary>
        public BorderStyle DefaultBorder { get; set; } = new BorderStyle();

        /// <summary>
        /// The default Date time format applied to cells which contains a <see cref="DateTime"/> value.
        /// </summary>
        public string DefaultDateTimeFormat { get; set; } = $"{CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern} hh:mm:ss";
    }
}
