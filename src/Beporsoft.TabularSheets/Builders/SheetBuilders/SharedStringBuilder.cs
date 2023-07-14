using Beporsoft.TabularSheets.Builders.Interfaces;
using Beporsoft.TabularSheets.Builders.StyleBuilders.SetupCollections;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Beporsoft.TabularSheets.Builders.SheetBuilders
{
    /// <summary>
    /// A class to handle the creation of the qualified node x:sharedStringTable of
    /// <see href="https://www.ecma-international.org/publications-and-standards/standards/ecma-376/">ECMA-376-1:2016 §18.4</see>
    /// <br/><br/>
    /// String values tipically aren't stored directly inside cells. Instead, string values are stored in a common, indexed list, which
    /// allows to store values only once. The cell which uses a shared string have a reference to the index of the string inside this table. A
    /// single SharedStringTable is required to one workbook, so it is shared across multiple sheets.
    /// </summary>
    internal class SharedStringBuilder
    {
        private readonly ISetupCollection<SharedStringSetup> _strings = new SharedStringSetupCollection();

        public int RegisterString(string str)
        {
            var stringSetup = new SharedStringSetup(str);
            int index = _strings.Register(stringSetup);
            return index;
        }

        public SharedStringTable GetSharedStringTable() => _strings.BuildContainer<SharedStringTable>();
        
    }
}
