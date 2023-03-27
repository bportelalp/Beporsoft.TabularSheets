using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Styling
{
    [DebuggerDisplay("Size={FontSize} | Color={FontColor} | Family={FontFamily}")]
    public class FontStyle : IEquatable<FontStyle?>
    {

        /// <summary>
        /// The font color, or <see langword="null"/> for default color, normally black.
        /// </summary>
        public Color? FontColor { get; set; } = null;

        /// <summary>
        /// The font size, or <see langword="null"/> for default size.
        /// </summary>
        public int? FontSize { get; set; } = null;

        /// <summary>
        /// The font family name, or <see langword="null"/> for default font
        /// </summary>
        public string? FontFamily { get; set; }


        internal static readonly FontStyle Default = new FontStyle();

        #region IEquatable
        public override bool Equals(object? obj)
        {
            return obj is FontStyle style && Equals(style);
        }

        public bool Equals(FontStyle? other)
        {
            return other is not null &&
                  EqualityComparer<Color?>.Default.Equals(FontColor, other.FontColor) &&
                  FontSize == other.FontSize &&
                  FontFamily == other.FontFamily;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(FontColor, FontSize, FontFamily);
        }


        #endregion
    }
}
