using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Options.ColumnWidth
{
    /// <summary>
    /// Represent an object which provide the column width calculation method for estimating the width
    /// </summary>
    public interface IColumnWidth
    {
        internal Builders.SheetBuilders.ContentMeasure InitializeContentMeasure();
    }
}
