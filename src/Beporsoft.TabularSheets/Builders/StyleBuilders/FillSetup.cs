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
    internal class FillSetup : Setup, IEquatable<FillSetup?>, IIndexedSetup
    {
        internal FillSetup(System.Drawing.Color foregroundColor, System.Drawing.Color? backgroundColor)
        {
            ForegroundColor = foregroundColor;
            BackgroundColor = backgroundColor;
        }

        public System.Drawing.Color ForegroundColor { get; set; }
        public System.Drawing.Color? BackgroundColor { get; set; }


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
                   BackgroundColor.Equals(other.BackgroundColor) &&
                   ForegroundColor.Equals(other.ForegroundColor);
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(BackgroundColor, ForegroundColor);
        }

        #region Build child elements
        private PatternFill BuildPatternFill()
        {
            var patternFill = new PatternFill
            {
                ForegroundColor = GetForegroundColor(),
                BackgroundColor = GetBackgroundColor(),
                PatternType = PatternValues.Solid
            };
            return patternFill;
        }

        private ForegroundColor GetForegroundColor()
        {
            return new ForegroundColor()
            {
                Rgb = OpenXMLHelpers.BuildHexBinaryFromColor(ForegroundColor),
            };
        }

        private BackgroundColor? GetBackgroundColor()
        {
            BackgroundColor? bg = null;
            if (BackgroundColor is not null)
            {
                bg = new BackgroundColor()
                {
                    Rgb = OpenXMLHelpers.BuildHexBinaryFromColor(BackgroundColor!.Value),
                };
            }
            return bg;
        }
        #endregion
    }
}
