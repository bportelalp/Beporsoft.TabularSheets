using Beporsoft.TabularSheets.Builders.Interfaces;
using Beporsoft.TabularSheets.Tools;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;

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
        public string FontName { get; set; } = null!;

        public override OpenXmlElement Build()
        {
            var font = new Font
            {
                FontSize = BuildFontSize(),
                FontName = BuildFontName(),
                Color = BuildColor()
            };
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
                   FontSize == other.FontSize &&
                   FontName == other.FontName;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(FontColor, FontSize, FontName);
        }

        #region Build Child properties
        private FontSize? BuildFontSize()
        {
            FontSize? fontSize = null;
            if (FontSize is not null)
            {
                fontSize = new FontSize() { Val = FontSize };
            }
            return fontSize;
        }
        private Color? BuildColor()
        {
            Color? color = null;
            if (FontColor is not null)
            {
                color = new Color() { Rgb = OpenXMLHelpers.BuildHexBinaryFromColor(FontColor.Value) };
            }
            return color;
        }
        private FontName? BuildFontName()
        {
            FontName? fontName = null;
            if (!string.IsNullOrWhiteSpace(FontName))
            {
                fontName = new FontName()
                {
                    Val = FontName
                };
            }
            return fontName;
        }
        #endregion
    }
}
