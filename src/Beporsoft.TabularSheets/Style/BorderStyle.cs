using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Style
{
    public class BorderStyle : IEquatable<BorderStyle?>
    {

        /// <summary>
        /// Border color for all sides
        /// </summary>
        public Color? Color { get; set; } = null;

        /// <summary>
        /// Border type for left border
        /// </summary>
        public BorderType Left { get; set; } = BorderType.None;

        /// <summary>
        /// Border type for right border
        /// </summary>
        public BorderType Right { get; set; } = BorderType.None;

        /// <summary>
        /// Border type for top border
        /// </summary>
        public BorderType Top { get; set; } = BorderType.None;

        /// <summary>
        /// Border type for bottom border
        /// </summary>
        public BorderType Bottom { get; set; } = BorderType.None;

        internal static BorderStyle Default { get; set; } = new BorderStyle();

        /// <summary>
        /// Set the same <see cref="BorderType"/> for all sides
        /// </summary>
        public void SetAll(BorderType borderType)
        {
            Left = borderType;
            Right = borderType;
            Top = borderType;
            Bottom = borderType;
        }

        public enum BorderType
        {
            None,
            Thin,
            Medium,
        }

        #region IEquatable
        public override bool Equals(object? obj)
        {
            return Equals(obj as BorderStyle);
        }

        public bool Equals(BorderStyle? other)
        {
            return other is not null &&
                   EqualityComparer<Color?>.Default.Equals(Color, other.Color) &&
                   Left == other.Left &&
                   Right == other.Right &&
                   Top == other.Top &&
                   Bottom == other.Bottom;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(Color, Left, Right, Top, Bottom);
        }
        #endregion
    }
}
