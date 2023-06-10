using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.CellStyling
{
    /// <summary>
    /// A container which represents all the settings to style cells
    /// </summary>
    public class Style : IEquatable<Style?>
    {
        /// <summary>
        /// <inheritdoc cref="FontStyle"/> Applies to the entire table
        /// </summary>
        public FontStyle Font { get; set; } = new();

        /// <summary>
        /// <inheritdoc cref="FillStyle"/> Applies to the entire table
        /// </summary>
        public FillStyle Fill { get; set; } = new();

        /// <summary>
        /// <inheritdoc cref="BorderStyle"/> Applies to the entire table
        /// </summary>
        public BorderStyle Border { get; set; } = new();

        /// <summary>
        /// Defines the numbering pattern applied to the entire table
        /// </summary>
        public string? NumberingPattern { get; set; } = null;

        internal static Style Default = new();

        #region IEquatable

        /// <inheritdoc cref="Equals(object?)"/>
        public override bool Equals(object? obj)
        {
            return Equals(obj as Style);
        }

        /// <inheritdoc cref="IEquatable{T}.Equals(T)"/>
        public bool Equals(Style? other)
        {
            return other is not null &&
                   EqualityComparer<FontStyle>.Default.Equals(Font, other.Font) &&
                   EqualityComparer<FillStyle>.Default.Equals(Fill, other.Fill) &&
                   EqualityComparer<BorderStyle>.Default.Equals(Border, other.Border);
        }

        /// <inheritdoc cref="GetHashCode"/>
        public override int GetHashCode()
        {
            return HashCode.Combine(Font, Fill, Border, NumberingPattern);
        }
        #endregion
    }
}
