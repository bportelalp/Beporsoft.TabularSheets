﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;

namespace Beporsoft.TabularSheets.CellStyling
{
    /// <summary>
    /// Defines the fill of cells.
    /// </summary>
    [DebuggerDisplay("Bg={BackgroundColor}")]
    public class FillStyle : IEquatable<FillStyle?>
    {
        /// <summary>
        /// Gets or sets the color for the background of the cell.
        /// </summary>
        public Color? BackgroundColor { get; set; }

        internal DocumentFormat.OpenXml.Spreadsheet.PatternValues PatternValue { get; set; } = DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid;

        internal static FillStyle Default { get; } = new FillStyle();

        #region IEquatable

        /// <inheritdoc cref="Equals(object?)"/>
        public override bool Equals(object? obj)
        {
            return Equals(obj as FillStyle);
        }

        /// <inheritdoc cref="IEquatable{T}.Equals(T)"/>
        public bool Equals(FillStyle? other)
        {
            return other is not null &&
                   EqualityComparer<Color?>.Default.Equals(BackgroundColor, other.BackgroundColor) &&
                   PatternValue.Equals(other.PatternValue);
        }

        /// <inheritdoc cref="GetHashCode"/>
        public override int GetHashCode()
        {
            return HashCode.Combine(BackgroundColor, PatternValue);
        }

        #endregion
    }
}
