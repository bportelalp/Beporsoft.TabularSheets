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
    internal class FillSetup : IEquatable<FillSetup?>, IStyleSetup
    {
        internal FillSetup(System.Drawing.Color foregroundColor, System.Drawing.Color? backgroundColor)
        {
            ForegroundColor = foregroundColor;
            BackgroundColor = backgroundColor;
        }

        public int Index { get; set; }
        public System.Drawing.Color ForegroundColor { get; init; }
        public System.Drawing.Color? BackgroundColor { get; init; }


        public OpenXmlElement Build()
        {
            var patternFill = new PatternFill();
            patternFill.ForegroundColor = GetForegroundColor();
            if (BackgroundColor is not null)
                patternFill.BackgroundColor = GetBackgroundColor();
            patternFill.PatternType = PatternValues.Solid;
            var fill = new Fill(patternFill);
            return fill;
        }

        public override bool Equals(object? obj)
        {
            return Equals(obj as FillSetup);
        }

        public bool Equals(FillSetup? other)
        {
            return other is not null &&
                   BackgroundColor.Equals(other.BackgroundColor);
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(BackgroundColor);
        }


        private ForegroundColor GetForegroundColor()
        {
            return new ForegroundColor()
            {
                Rgb = OpenXMLHelpers.BuildHexBinaryFromColor(ForegroundColor),
            };
        }
        private BackgroundColor GetBackgroundColor()
        {
            return new BackgroundColor()
            {
                Rgb = OpenXMLHelpers.BuildHexBinaryFromColor(BackgroundColor!.Value),
            };
        }
    }
}
