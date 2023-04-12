using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.CellStyling
{
    /// <summary>
    /// Defines the fill of cells
    /// </summary>
    [DebuggerDisplay("Bg={BackgroundColor}")]
    public class FillStyle : IEquatable<FillStyle?>
    {

        public Color? BackgroundColor { get; set; }

        internal DocumentFormat.OpenXml.Spreadsheet.PatternValues PatternValue { get; set; } = DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid;

        internal static FillStyle Default { get; } = new FillStyle();

        #region IEquatable

        public override bool Equals(object? obj)
        {
            return Equals(obj as FillStyle);
        }

        public bool Equals(FillStyle? other)
        {
            return other is not null &&
                   EqualityComparer<Color?>.Default.Equals(BackgroundColor, other.BackgroundColor) &&
                   PatternValue.Equals(other.PatternValue);
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(BackgroundColor, PatternValue);
        }

        #endregion
    }
}
