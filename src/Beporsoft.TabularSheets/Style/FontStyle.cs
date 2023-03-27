using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Style
{
    [DebuggerDisplay("Size={FontSize} | Color={FontColor} | Family={FontFamily}")]
    public class FontStyle : IEquatable<FontStyle?>
    {
        internal static readonly FontStyle Default = new FontStyle();

        public Color? FontColor { get; set; } = null;
        public int? FontSize { get; set; } = null;
        public string? FontFamily { get; set; }

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
