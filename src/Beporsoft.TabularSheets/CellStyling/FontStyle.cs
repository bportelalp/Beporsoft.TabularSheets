using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.CellStyling
{
    /// <summary>
    /// Defines the font of cells.
    /// </summary>
    [DebuggerDisplay("Size={Size} | Color={Color} | Font={Font}")]
    public class FontStyle : IEquatable<FontStyle?>
    {

        /// <summary>
        /// The font color, or <see langword="null"/> for default color, normally black.
        /// </summary>
        public Color? Color { get; set; } = null;

        /// <summary>
        /// The font size, or <see langword="null"/> for default size.
        /// </summary>
        public int? Size { get; set; } = null;

        /// <summary>
        /// The font family name, or <see langword="null"/> for default font
        /// </summary>
        public string? Font { get; set; }

        internal static FontStyle Default { get; } = new FontStyle();

        #region IEquatable

        /// <inheritdoc cref="Equals(object?)"/>
        public override bool Equals(object? obj)
        {
            return obj is FontStyle style && Equals(style);
        }

        /// <inheritdoc cref="IEquatable{T}.Equals(T)"/>
        public bool Equals(FontStyle? other)
        {
            return other is not null &&
                  EqualityComparer<Color?>.Default.Equals(Color, other.Color) &&
                  Size == other.Size &&
                  Font == other.Font;
        }

        /// <inheritdoc cref="GetHashCode"/>
        public override int GetHashCode()
        {
            return HashCode.Combine(Color, Size, Font);
        }


        #endregion
    }
}
