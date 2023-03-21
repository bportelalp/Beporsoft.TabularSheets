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
    internal class IndexableFormat : IEquatable<IndexableFormat?>, IIndexableStyle
    {
        public IndexableFormat(IndexableFill? fill, IndexableFont? font)
        {
            Fill = fill;
            Font = font;
        }

        public int Index { get; set; }
        public IndexableFill? Fill { get; init; }
        public IndexableFont? Font { get; init; }

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
            return Equals(obj as IndexableFormat);
        }

        public bool Equals(IndexableFormat? other)
        {
            return other is not null &&
                   EqualityComparer<IndexableFill?>.Default.Equals(Fill, other.Fill) &&
                   EqualityComparer<IndexableFont?>.Default.Equals(Font, other.Font);
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(Fill, Font);
        }
    }
}
