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
    /// Builder for the qualified node x:fill of 
    /// <see href="https://www.ecma-international.org/publications-and-standards/standards/ecma-376/">ECMA-376-1:2016 §18.8.20</see>
    /// </summary>
    [DebuggerDisplay("Id={Index}")]
    internal class FillSetup : Setup, IEquatable<FillSetup?>, IIndexedSetup
    {
        internal FillSetup(FillStyle fill)
        {
            Fill = fill;
        }

        public FillStyle Fill { get; set; }


        public override OpenXmlElement Build()
        {
            var fill = new Fill()
            {
                PatternFill = BuildPatternFill()
            };
            return fill;
        }

        public static FillSetup FromOpenXmlFill(Fill fillXml)
        {
            FillStyle fill = new()
            {
                BackgroundColor = fillXml.PatternFill?.ForegroundColor?.Rgb?.FromOpenXmlHexBinaryValue(),
                PatternValue = fillXml.PatternFill!.PatternType!.Value
            };
            return new FillSetup(fill);
        }

        #region IEquatable
        public override bool Equals(object? obj)
        {
            return Equals(obj as FillSetup);
        }

        public bool Equals(FillSetup? other)
        {
            return other is not null &&
                EqualityComparer<FillStyle>.Default.Equals(Fill, other.Fill);
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(Fill);
        }
        #endregion

        #region Build child elements
        private PatternFill BuildPatternFill()
        {
            PatternFill patternFill = new PatternFill
            {
                ForegroundColor = BuildForegroundColor(),
                PatternType = Fill.PatternValue
            };
            return patternFill;
        }


        private ForegroundColor? BuildForegroundColor()
        {
            ForegroundColor? bg = null;
            if (Fill.BackgroundColor is not null)
            {

            }
            if (Fill.BackgroundColor is not null)
            {
                bg = new ForegroundColor()
                {
                    Rgb = Fill.BackgroundColor!.Value.ToOpenXmlHexBinary(),
                };
            }
            return bg;
        }
        #endregion
    }
}
