using System;
using System.Diagnostics;

namespace Beporsoft.TabularSheets.CellStyling
{
    /// <summary>
    /// Defines the alignment of the content inside the cell
    /// </summary>
    [DebuggerDisplay("Horizontal={Horizontal?.ToString()} | Vertical={Vertical?.ToString()} | TextWrap={TextWrap}")]
    public class AlignmentStyle : IEquatable<AlignmentStyle?>
    {
        /// <summary>
        /// Set to <see langword="true"/> for line-wrapped text within the cell
        /// </summary>
        public bool? TextWrap { get; set; }

        /// <summary>
        /// Horizontal alignment in cells
        /// </summary>
        public HorizontalAlignment? Horizontal { get; set; }

        /// <summary>
        /// Vertical alignment in cells
        /// </summary>
        public VerticalAlignment? Vertical { get; set; }

        internal static AlignmentStyle Default = new AlignmentStyle();


        #region IEquatable

        /// <inheritdoc cref="object.Equals(object?)"/>
        public override bool Equals(object? obj)
        {
            return Equals(obj as AlignmentStyle);
        }

        /// <inheritdoc cref="IEquatable{T}.Equals(T)"/>
        public bool Equals(AlignmentStyle? other)
        {
            return other is not null &&
                   TextWrap == other.TextWrap &&
                   Horizontal == other.Horizontal &&
                   Vertical == other.Vertical;
        }

        /// <inheritdoc cref="GetHashCode"/>
        public override int GetHashCode()
        {
            return HashCode.Combine(TextWrap, Horizontal, Vertical);
        }
        #endregion

        /// <summary>
        /// Possible values for horizontal alignment
        /// </summary>
        public enum HorizontalAlignment
        {
            /// <summary>
            /// Centered horizontal aligment. Text is centered across the cell
            /// </summary>
            Center,
            /// <summary>
            /// Text data is left-aligned. <br/>
            /// Numbers, dates and times are right-aligned. <br/>
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

        /// <summary>
        /// Possible values for vertical alignment
        /// </summary>
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
