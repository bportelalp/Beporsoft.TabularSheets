using Beporsoft.TabularSheets.Builders.Interfaces;
using Beporsoft.TabularSheets.Styling;
using Beporsoft.TabularSheets.Tools;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace Beporsoft.TabularSheets.Builders.StyleBuilders
{
    [DebuggerDisplay("Id={Index}")]
    internal class FontSetup : Setup, IEquatable<FontSetup?>, IIndexedSetup
    {

        internal FontSetup(FontStyle fontStyle)
        {
            FontStyle = fontStyle;
        }

        public FontStyle FontStyle { get; set; }

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


        #region Build Child properties
        private FontSize? BuildFontSize()
        {
            FontSize? fontSize = null;
            if (FontStyle.FontSize is not null)
            {
                fontSize = new FontSize() { Val = FontStyle.FontSize };
            }
            return fontSize;
        }
        private Color? BuildColor()
        {
            Color? color = null;
            if (FontStyle.FontColor is not null)
            {
                color = new Color() { Rgb = OpenXmlHelpers.BuildHexBinaryFromColor(FontStyle.FontColor.Value) };
            }
            return color;
        }
        private FontName? BuildFontName()
        {
            FontName? fontName = null;
            if (!string.IsNullOrWhiteSpace(FontStyle.FontFamily))
            {
                fontName = new FontName()
                {
                    Val = FontStyle.FontFamily
                };
            }
            return fontName;
        }

        public override bool Equals(object? obj)
        {
            return Equals(obj as FontSetup);
        }

        public bool Equals(FontSetup? other)
        {
            return other is not null  &&
                   EqualityComparer<FontStyle>.Default.Equals(FontStyle, other.FontStyle);
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(FontStyle);
        }



        #endregion
    }
}
