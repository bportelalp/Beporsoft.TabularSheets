using Beporsoft.TabularSheets.Builders.Shared;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Builders.SheetBuilders
{
    internal class SharedStringBuilder
    {
        private readonly IndexedSetupCollection<SharedStringSetup> _strings = new();

        public int RegisterString(string str)
        {
            var stringSetup = new SharedStringSetup(str);
            int index = _strings.Register(stringSetup);
            return index;
        }

        public SharedStringTable GetSharedStringTable() => _strings.BuildContainer<SharedStringTable>();
        
    }
}
