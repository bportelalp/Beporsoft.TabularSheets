using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;

namespace Beporsoft.TabularSheets.CellStyling
{
    /// <summary>
    /// Defines the cell border style.
    /// </summary>
    [DebuggerDisplay("Color={Color} | Borders=[{Top}, {Right}, {Bottom}, {Left}]")]
    public class BorderStyle : IEquatable<BorderStyle?>
    {

        /// <summary>
        /// Border color for all sides
        /// </summary>
        public Color? Color { get; set; } = null;

        /// <summary>
        /// Border type for left border
        /// </summary>
        public BorderType? Left { get; set; }

        /// <summary>
        /// Border type for right border
        /// </summary>
        public BorderType? Right { get; set; }

        /// <summary>
        /// Border type for top border
        /// </summary>
        public BorderType? Top { get; set; }

        /// <summary>
        /// Border type for bottom border
        /// </summary>
        public BorderType? Bottom { get; set; }

        internal static BorderStyle Default { get; } = new BorderStyle();

        #region SetBorder
        /// <summary>
        /// Set the border type to all sides
        /// </summary>
        /// <param name="borderType"></param>
        public void SetBorderType(BorderType? borderType)
        {
            Left = borderType;
            Right = borderType;
            Top = borderType;
            Bottom = borderType;
        }

        /// <summary>
        /// Set the border type for each side
        /// </summary>
        /// <param name="top"></param>
        /// <param name="right"></param>
        /// <param name="bottom"></param>
        /// <param name="left"></param>
        public void SetBorderType(BorderType? top, BorderType? right, BorderType? bottom, BorderType? left)
        {
            Top = top;
            Bottom = bottom;
            Left = left;
            Right = right;
        }

        /// <summary>
        /// Set the border type for horizontal and vertical sides
        /// </summary>
        /// <param name="horizontal"></param>
        /// <param name="vertical"></param>
        public void SetBorderType(BorderType? horizontal, BorderType? vertical)
        {
            Top = horizontal;
            Bottom = horizontal;
            Left = vertical;
            Right = vertical;
        }
        #endregion

        #region IEquatable
        /// <inheritdoc cref="Equals(object)"/>
        public override bool Equals(object? obj)
        {
            return Equals(obj as BorderStyle);
        }

        /// <inheritdoc cref="IEquatable{T}.Equals(T)"/>
        public bool Equals(BorderStyle? other)
        {
            return other is not null &&
                   EqualityComparer<Color?>.Default.Equals(Color, other.Color) &&
                   Left == other.Left &&
                   Right == other.Right &&
                   Top == other.Top &&
                   Bottom == other.Bottom;
        }


        /// <inheritdoc cref="GetHashCode"/>
        public override int GetHashCode()
        {
            return HashCode.Combine(Color, Left, Right, Top, Bottom);
        }
        #endregion


        /// <summary>
        /// Type of border
        /// </summary>
        public enum BorderType
        {
            /// <summary>
            /// No border is present
            /// </summary>
            None,
            /// <summary>
            /// Border thin
            /// </summary>
            Thin,
            /// <summary>
            /// Border medium
            /// </summary>
            Medium,
            /// <summary>
            /// Border thick
            /// </summary>
            Thick,
            /// <summary>
            /// Border dashed
            /// </summary>
            Dashed,
            /// <summary>
            /// Border dotted
            /// </summary>
            Dotted,
            /// <summary>
            /// Border is dash-dot
            /// </summary>
            DashDot,
            /// <summary>
            /// Border is dash-dot-dot
            /// </summary>
            DashDotDot,
            /// <summary>
            /// Border is hairline
            /// </summary>
            Hair,
            /// <summary>
            /// Border is medium dash-dot
            /// </summary>
            MediumDashDot,
            /// <summary>
            /// Border is medium dash-dot-dot
            /// </summary>
            MediumDashDotDot,
            /// <summary>
            /// Border is medium dashed
            /// </summary>
            MediumDashed,
            /// <summary>
            /// Border is slant-dash-dot
            /// </summary>
            SlantDashDot,
        }
    }
}
