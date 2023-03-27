using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Style
{
    /// <summary>
    /// Defines the style of every cell of the heading of <see cref="TabularSheet{T}"/> inside a 
    /// Spreadsheet document.<br/> Initially all styles are the default ones.
    /// </summary>
    public class HeaderStyle
    {
        /// <summary>
        /// The <see cref="FillStyle"/> applied to header of <see cref="TabularSheet{T}"/>.
        /// </summary>
        public FillStyle Fill { get; set; } = new FillStyle();


        /// <summary>
        /// The <see cref="FontStyle"/> applied to header of <see cref="TabularSheet{T}"/>.
        /// </summary>
        public FontStyle Font { get; set; } = new FontStyle();


        /// <summary>
        /// The <see cref="BorderStyle"/> applied to header of <see cref="TabularSheet{T}"/>.
        /// </summary>
        public BorderStyle Border { get; set; } = new BorderStyle();

    }
}
