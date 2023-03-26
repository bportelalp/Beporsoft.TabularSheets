using Beporsoft.TabularSheets.Builders.Shared;
using Beporsoft.TabularSheets.Tools;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Builders.StyleBuilders
{
    internal class FontSetup : Setup, IEquatable<FontSetup?>, IIndexedSetup
    {

        internal FontSetup(System.Drawing.Color fontColor, int fontSize)
        {
            FontColor = fontColor;
            FontSize = fontSize;
        }

        public System.Drawing.Color? FontColor { get; set; }
        public int? FontSize { get; set; } = 10;

        public override OpenXmlElement Build()
        {
            var font = new Font();

            if (FontSize is not null)
                font.FontSize = BuildFontSize();
            if(FontColor is not null)
                font.Color = BuildColor();

            return font;
        }

        public override bool Equals(object? obj)
        {
            return Equals(obj as FontSetup);
        }

        public bool Equals(FontSetup? other)
        {
            return other is not null &&
                   EqualityComparer<System.Drawing.Color?>.Default.Equals(FontColor, other.FontColor) &&
                   FontSize == other.FontSize;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(FontColor, FontSize);
        }

        private FontSize BuildFontSize() => new FontSize() { Val = FontSize };
        private Color BuildColor() => new Color() { Rgb = OpenXMLHelpers.BuildHexBinaryFromColor(FontColor!.Value) };
    }
}
