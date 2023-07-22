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
        /// Set the width application method for the calculation of the column width
        /// </summary>
        public IColumnWidth? Width { get; set; }
    }
}
