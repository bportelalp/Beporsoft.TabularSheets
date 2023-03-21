using Beporsoft.TabularSheets.Builders.Interfaces;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Builders.StyleBuilders
{
    internal class IndexableFill : IEquatable<IndexableFill?>, IIndexableStyle
    {
        public IndexableFill(System.Drawing.Color backgroundColor)
        {
            BackgroundColor = backgroundColor;

        }
        public System.Drawing.Color BackgroundColor { get; }


        public OpenXmlElement Build()
        {
            var hex = $"FF{BackgroundColor.R:X}{BackgroundColor.G:X}{BackgroundColor.B:X}";
            var bgCol = new ForegroundColor()
            {
                Rgb = new HexBinaryValue(hex) //TODO-Fill the color
            };
            var patternFill = new PatternFill(bgCol);
            patternFill.PatternType = PatternValues.Solid;
            var fill = new Fill(patternFill);
            return fill;
        }

        public override bool Equals(object? obj)
        {
            return Equals(obj as IndexableFill);
        }

        public bool Equals(IndexableFill? other)
        {
            return other is not null &&
                   BackgroundColor.Equals(other.BackgroundColor);
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(BackgroundColor);
        }
    }
}
