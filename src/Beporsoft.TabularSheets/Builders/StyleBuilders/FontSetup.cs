using Beporsoft.TabularSheets.Builders.Interfaces;
using Beporsoft.TabularSheets.CellStyling;
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
            if (FontStyle.Size is not null)
            {
                fontSize = new FontSize() { Val = FontStyle.Size };
            }
            return fontSize;
        }
        private Color? BuildColor()
        {
            Color? color = null;
            if (FontStyle.Color is not null)
            {
                color = new Color() { Rgb = FontStyle.Color.Value.ToHexBinaryOpenXml() };
            }
            return color;
        }
        private FontName? BuildFontName()
        {
            FontName? fontName = null;
            if (!string.IsNullOrWhiteSpace(FontStyle.Font))
            {
                fontName = new FontName()
                {
                    Val = FontStyle.Font
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
