using Beporsoft.TabularSheets.Builders.Interfaces;
using Beporsoft.TabularSheets.Style;
using Beporsoft.TabularSheets.Tools;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace Beporsoft.TabularSheets.Builders.StyleBuilders
{
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

        #region Build child elements
        private PatternFill BuildPatternFill()
        {
            var patternFill = new PatternFill
            {
                ForegroundColor = BuildForegroundColor(),
                PatternType = PatternValues.Solid
            };
            return patternFill;
        }


        private ForegroundColor? BuildForegroundColor()
        {
            ForegroundColor? bg = null;
            if(Fill.BackgroundColor is not null)
            {

            }
            if (Fill.BackgroundColor is not null)
            {
                bg = new ForegroundColor()
                {
                    Rgb = OpenXMLHelpers.BuildHexBinaryFromColor(Fill.BackgroundColor!.Value),
                };
            }
            return bg;
        }
        #endregion
    }
}
