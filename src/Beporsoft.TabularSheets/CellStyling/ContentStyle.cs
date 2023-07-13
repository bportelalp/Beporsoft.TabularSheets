using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.CellStyling
{
    public class ContentStyle
    {
        public bool? TextWrap { get; set; }

        public HorizontalAlignment? HorizontalAlign { get; set; }

        public VerticalAlignment? VerticalAlign { get; set; }

        public enum HorizontalAlignment
        {
            /// <summary>
            /// Centered horizontal aligment. Text is centered across the cell
            /// </summary>
            Center,
            /// <summary>
            /// Text data is left-aligned.<br/>
            /// Numbers, dates and times are right-aligned.<br/>
            /// Booleans are centered
            /// </summary>
            General,
            /// <summary>
            /// Text justify
            /// </summary>
            Justify,
            /// <summary>
            /// Left-aligned
            /// </summary>
            Left,
            /// <summary>
            /// Right-aligned
            /// </summary>
            Right,
        }

        public enum VerticalAlignment
        {
            /// <summary>
            /// Align to bottom
            /// </summary>
            Bottom,
            /// <summary>
            /// Align to center across the height
            /// </summary>
            Center,
            /// <summary>
            /// Align to top
            /// </summary>
            Top,
        }
    }
}
