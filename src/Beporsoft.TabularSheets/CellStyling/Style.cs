using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Styling
{
    public class Style : IEquatable<Style?>
    {
        public FontStyle Font { get; set; } = new();
        public FillStyle Fill { get; set; } = new();
        public BorderStyle Border { get; set; } = new();

        internal static Style Default = new();

        #region IEquatable
        public override bool Equals(object? obj)
        {
            return Equals(obj as Style);
        }

        public bool Equals(Style? other)
        {
            return other is not null &&
                   EqualityComparer<FontStyle>.Default.Equals(Font, other.Font) &&
                   EqualityComparer<FillStyle>.Default.Equals(Fill, other.Fill) &&
                   EqualityComparer<BorderStyle>.Default.Equals(Border, other.Border);
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(Font, Fill, Border);
        }
        #endregion
    }
}
