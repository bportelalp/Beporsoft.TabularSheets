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
    /// Builder for the qualified node x:border of 
    /// <see href="https://www.ecma-international.org/publications-and-standards/standards/ecma-376/">ECMA-376-1:2016 §18.8.4</see>
    /// </summary>
    [DebuggerDisplay("Id={Index}")]
    internal class BorderSetup : Setup, IIndexedSetup, IEquatable<BorderSetup?>
    {
        internal BorderSetup(BorderStyle borderStyle)
        {
            BorderStyle = borderStyle;
        }

        public BorderStyle BorderStyle { get; set; }

        public override OpenXmlElement Build()
        {
            Border border = new Border()
            {
                LeftBorder = BuildBorder<LeftBorder>(BorderStyle.Left),
                RightBorder = BuildBorder<RightBorder>(BorderStyle.Right),
                TopBorder = BuildBorder<TopBorder>(BorderStyle.Top),
                BottomBorder = BuildBorder<BottomBorder>(BorderStyle.Bottom),
            };
            return border;
        }

        #region IEquatable
        public override bool Equals(object? obj)
        {
            return Equals(obj as BorderSetup);
        }

        public bool Equals(BorderSetup? other)
        {
            return other is not null &&
                EqualityComparer<BorderStyle>.Default.Equals(BorderStyle, other.BorderStyle);
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(BorderStyle);
        }
        #endregion

        #region Build Child Elements
        private T? BuildBorder<T>(BorderStyle.BorderType borderType) where T : BorderPropertiesType, new()
        {
            T? border = null;
            if (borderType is not BorderStyle.BorderType.None)
            {
                bool parsedOk = Enum.TryParse(borderType.ToString(), out BorderStyleValues style);
                if (parsedOk)
                {
                    border = new T()
                    {
                        Style = style,
                    };
                    if (BorderStyle.Color is not null)
                    {
                        border.Color = new Color()
                        {
                            Rgb = BorderStyle.Color.Value.ToOpenXmlHexBinary()
                        };
                    }
                }
            }
            return border;
        }
        #endregion
    }
}
