using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Beporsoft.TabularSheets.Builders.Interfaces
{
    /// <summary>
    /// Provides an abstractions of nodes which must be indexed inside a parent container to assign
    /// the respective index to <see cref="Cell"/> properties. Examples: <br/> 
    /// - SharedStrings: <see cref="CellType.CellValue"/> must be the index inside <see cref="SharedStringTable"/>.<br/>
    /// - Styles: <see cref="CellType.StyleIndex"/> must be the index inside <see cref="Stylesheet.CellFormats"/>.
    /// </summary>
    internal interface IIndexedSetup
    {

        /// <summary>
        /// The index inside the collection stored on the OpenXml Tree.
        /// </summary>
        int Index { get; }

        /// <summary>
        /// Generates the <see cref="OpenXmlElement"/> to fill the tree.
        /// </summary>
        OpenXmlElement Build();
    }
}
