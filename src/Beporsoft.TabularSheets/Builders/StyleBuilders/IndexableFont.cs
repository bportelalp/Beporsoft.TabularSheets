using Beporsoft.TabularSheets.Builders.Interfaces;
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
    internal class IndexableFont : IEquatable<IndexableFont?>, IIndexableStyle
    {

        public IndexableFont(System.Drawing.Color fontColor, int fontSize)
        {
            FontColor = fontColor;
            FontSize = fontSize;
        }

        public int Index { get; set; }
        public System.Drawing.Color? FontColor { get; init; }
        public int? FontSize { get; init; } = 10;

        public OpenXmlElement Build()
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
            return Equals(obj as IndexableFont);
        }

        public bool Equals(IndexableFont? other)
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
