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
    internal class FormatSetup : IEquatable<FormatSetup?>, IStyleSetup
    {
        public FormatSetup(FillSetup? fill, FontSetup? font)
        {
            Fill = fill;
            Font = font;
        }

        public int Index { get; set; }
        public FillSetup? Fill { get; set; }
        public FontSetup? Font { get; set; }

        public OpenXmlElement Build()
        {
            var format = new CellFormat();
            if (Font is not null)
                format.FontId = new UInt32Value(Convert.ToUInt32(Font.Index));
            if (Fill is not null)
                format.FillId = new UInt32Value(Convert.ToUInt32(Fill.Index));
            return format;
        }

        public override bool Equals(object? obj)
        {
            return Equals(obj as FormatSetup);
        }

        public bool Equals(FormatSetup? other)
        {
            return other is not null &&
                   EqualityComparer<FillSetup?>.Default.Equals(Fill, other.Fill) &&
                   EqualityComparer<FontSetup?>.Default.Equals(Font, other.Font);
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(Index, Fill, Font);
        }
    }
}
