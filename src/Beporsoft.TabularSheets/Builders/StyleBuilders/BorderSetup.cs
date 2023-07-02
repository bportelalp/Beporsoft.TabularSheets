using Beporsoft.TabularSheets.Builders.Interfaces;
using Beporsoft.TabularSheets.CellStyling;
using Beporsoft.TabularSheets.Tools;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

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

        public static BorderSetup FromOpenXmlBorder(Border borderXml)
        {
            BorderStyle border = new();
            List<System.Drawing.Color> colorsBorders = new List<System.Drawing.Color>();
            BorderStyle.BorderType parsedBorder;
            bool parsedOk = false;

            // Left border
            BorderStyleValues? leftBorderXml = borderXml.LeftBorder?.Style?.Value;
            parsedOk = Enum.TryParse(leftBorderXml?.ToString(), out parsedBorder);
            border.Left = parsedOk ? parsedBorder : null;

            if (borderXml.LeftBorder?.Color?.Rgb is not null)
                colorsBorders.Add(borderXml.LeftBorder.Color.Rgb.FromOpenXmlHexBinaryValue());

            // Right border
            BorderStyleValues? rightBorderXml = borderXml.RightBorder?.Style?.Value;
            parsedOk = Enum.TryParse(rightBorderXml?.ToString(), out parsedBorder);
            border.Right = parsedOk ? parsedBorder : null;
            if (borderXml.RightBorder?.Color?.Rgb is not null)
                colorsBorders.Add(borderXml.RightBorder.Color.Rgb.FromOpenXmlHexBinaryValue());

            // Top border
            BorderStyleValues? topBorderXml = borderXml.TopBorder?.Style?.Value;
            parsedOk = Enum.TryParse(topBorderXml?.ToString(), out parsedBorder);
            border.Top = parsedOk ? parsedBorder : null;
            if (borderXml.TopBorder?.Color?.Rgb is not null)
                colorsBorders.Add(borderXml.TopBorder.Color.Rgb.FromOpenXmlHexBinaryValue());

            // Bottom border
            BorderStyleValues? bottomBorderXml = borderXml.BottomBorder?.Style?.Value;
            parsedOk = Enum.TryParse(bottomBorderXml?.ToString(), out parsedBorder);
            border.Bottom = parsedOk ? parsedBorder : null;
            if (borderXml.BottomBorder?.Color?.Rgb is not null)
                colorsBorders.Add(borderXml.BottomBorder.Color.Rgb.FromOpenXmlHexBinaryValue());

            // Take the color more frequent
            var colorGrouped = colorsBorders.GroupBy(c => c.ToArgb()).OrderBy(c => c.Count());
            border.Color = colorGrouped.FirstOrDefault()?.FirstOrDefault();
            return new BorderSetup(border);
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
        private T? BuildBorder<T>(BorderStyle.BorderType? borderType) where T : BorderPropertiesType, new()
        {
            T? border = null;
            if (borderType is not null || borderType is not BorderStyle.BorderType.None)
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
