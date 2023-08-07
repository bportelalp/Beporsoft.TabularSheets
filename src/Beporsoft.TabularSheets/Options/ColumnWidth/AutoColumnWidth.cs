using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Options.ColumnWidth
{
    /// <summary>
    /// Implements the auto calculation of the column width based on the content of the column.
    /// This method fits the width to the content for Calibri at any size, or similar fonts.
    /// For other fonts, use <see cref="ScaleFactor"/> to perform better.
    /// </summary>
    public class AutoColumnWidth : IColumnWidth
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        public AutoColumnWidth()
        {

        }

        /// <summary>
        /// A scale factor to improve the performance of the method for fonts different of Calibri.
        /// </summary>
        /// <param name="scaleFactor">A factor applied for fonts different of calibri</param>
        public AutoColumnWidth(double scaleFactor)
        {
            ScaleFactor = scaleFactor;
        }

        /// <summary>
        /// Scale factor applied to the result of the calculation
        /// </summary>
        public double ScaleFactor { get; set; } = 1.0;
    }
}
