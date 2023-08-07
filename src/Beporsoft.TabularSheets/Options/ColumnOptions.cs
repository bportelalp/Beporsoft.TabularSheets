using Beporsoft.TabularSheets.Options.ColumnWidth;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Options
{
    /// <summary>
    /// Common options for column configuration
    /// </summary>
    public class ColumnOptions
    {
        /// <summary>
        /// Set the width application method for the calculation of the column width.<br></br><br></br>
        /// Use <see cref="FixedColumnWidth"/> for a fixed value or <see cref="AutoColumnWidth"/> for automatic
        /// width calculation based on content
        /// </summary>
        public IColumnWidth? Width { get; set; }
    }
}
