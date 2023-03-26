using Beporsoft.TabularSheets.Builders.Interfaces;
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
    internal class FormatSetup : Setup, IEquatable<FormatSetup?>, IIndexedSetup
    {
        public FormatSetup(FillSetup? fill, FontSetup? font, BorderSetup? border)
        {
            Fill = fill;
            Font = font;
            Border = border;
        }

        public FillSetup? Fill { get; set; }
        public FontSetup? Font { get; set; }
        public BorderSetup? Border { get; set; }

        public override OpenXmlElement Build()
        {
            var format = new CellFormat();
            if (Font is not null)
                format.FontId = OpenXMLHelpers.ToUint32Value(Font.Index);
            if (Fill is not null)
                format.FillId = OpenXMLHelpers.ToUint32Value(Fill.Index);
            if (Border is not null)
                format.BorderId = OpenXMLHelpers.ToUint32Value(Border.Index);
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
                   EqualityComparer<FontSetup?>.Default.Equals(Font, other.Font) &&
                   EqualityComparer<BorderSetup?>.Default.Equals(Border, other.Border); ;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(Fill, Font, Border);
        }
    }
}
