using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Builders.SheetBuilders
{
    /// <summary>
    /// Container of the measurement of content
    /// </summary>
    internal class ContentMeasure
    {
        /// <summary>
        /// If <see langword="true"/>, <see cref="MaxContentWidth"/> will be updating with new values.
        /// Otherwise, <see cref="MaxContentWidth"/> is the default value on the initialization and it doesn't change.
        /// </summary>
        public bool AutoWidth { get; set; }

        /// <summary>
        /// Value of width to apply to the column
        /// </summary>
        public double MaxContentWidth { get; set; }

        /// <summary>
        /// Max font size applied to any cell column
        /// </summary>
        public double MaxFontSize { get; set; } = BuildHelpers.DefaultFontSize;

        /// <summary>
        /// A scale applied for fonts different from calibri
        /// </summary>
        public double FontFactor { get; set; } = 1.0;

        /// <summary>
        /// Check a new value and save if it is greather than the current.
        /// </summary>
        /// <param name="width"></param>
        /// <param name="fontSize"></param>
        public void EvaluateWidth(double width, double? fontSize)
        {
            if (!AutoWidth)
                return;
            if (MaxContentWidth < width)
                MaxContentWidth = width;
            if (fontSize is not null && fontSize > MaxFontSize)
                MaxFontSize = fontSize.Value;
        }
    }
}
