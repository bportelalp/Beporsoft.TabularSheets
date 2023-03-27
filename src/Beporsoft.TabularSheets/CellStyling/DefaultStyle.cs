using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Styling
{
    /// <summary>
    /// Defines the default styles for every cell of the representation of <see cref="TabularSheet{T}"/> inside
    /// a Spreadsheet Document, if there is no more specific one.<br/> Initially all the styles are the default ones
    /// </summary>
    public class DefaultStyle : Style
    {

        /// <summary>
        /// The default Date time format applied to cells which contains a <see cref="DateTime"/> value.
        /// </summary>
        public string DefaultDateTimeFormat { get; set; } = $"{CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern} hh:mm:ss";

    }
}
