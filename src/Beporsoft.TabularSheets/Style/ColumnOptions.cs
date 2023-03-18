using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Style
{
    public class ColumnOptions
    {
        public const double DefaultWidth = 10.71;

        /// <summary>
        /// Width in dpi of column
        /// </summary>
        public double Width { get; set; } = DefaultWidth;
    }
}
