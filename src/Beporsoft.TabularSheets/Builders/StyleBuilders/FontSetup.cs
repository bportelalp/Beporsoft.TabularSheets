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
    /// <summary>
    /// Builder for the qualified node x:font of 
    /// <see href="https://www.ecma-international.org/publications-and-standards/standards/ecma-376/">ECMA-376-1:2016 §18.8.22</see>
    /// </summary>
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
                Color = BuildColor(),
                Bold = BuildBold(),
                Italic = BuildItalic(),
            };
            return font;
        }

        public static FontSetup FromOpenXmlFont(Font fontXml)
        {
            FontStyle font = new FontStyle()
            {
                Font = fontXml.FontName?.Val?.Value,
                Color = fontXml.Color?.Rgb?.FromOpenXmlHexBinaryValue(),
                Size = fontXml.FontSize?.Val?.Value,
            };
            return new FontSetup(font);
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
                color = new Color() { Rgb = FontStyle.Color.Value.ToOpenXmlHexBinary() };
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

        private Bold? BuildBold()
        {
            Bold? bold = null;
            if(FontStyle.Bold is not null)
            {
                bold = new Bold()
                {
                    Val = FontStyle.Bold
                };
            }
            return bold;
        }

        private Italic? BuildItalic()
        {
            Italic? italic = null;
            if (FontStyle.Italic is not null)
            {
                italic = new Italic()
                {
                    Val = FontStyle.Bold
                };
            }
            return italic;
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
