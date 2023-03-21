using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Builders.Interfaces
{
    /// <summary>
    /// Provides an abstractions of generate syles configurations which must be indexed inside the stylesheet to assign
    /// correctly to cells.
    /// </summary>
    internal interface IIndexableStyle
    {
        int Index { get; set; }

        /// <summary>
        /// Generates the <see cref="OpenXmlElement"/> to fill the Stylesheet tree.
        /// </summary>
        /// <returns></returns>
        OpenXmlElement Build();
    }
}
