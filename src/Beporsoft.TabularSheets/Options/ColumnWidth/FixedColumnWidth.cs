using Beporsoft.TabularSheets.Builders.SheetBuilders;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Options.ColumnWidth
{
    /// <summary>
    /// Represent a fixed width for column as the number of numeric digits which fit on a cell for Calibri font or similar,
    /// at any size.
    /// </summary>
    public class FixedColumnWidth : IColumnWidth
    {
        /// <summary>
        /// Fixed width for column of <paramref name="width"/> as the number of numeric digits which fit on cell, for Calibri
        /// </summary>
        /// <param name="width"></param>
        public FixedColumnWidth(double width)
        {
            Width = width;
        }

        /// <summary>
        /// Width as the number of numeric digits which fit on cell, for Calibri 11.
        /// </summary>
        public double Width { get; set; }

        ContentMeasure IColumnWidth.InitializeContentMeasure()
        {
            ContentMeasure measure = new()
            {
                MaxContentWidth = Width,
                AutoWidth = false
            };
            return measure;
        }
    }
}
