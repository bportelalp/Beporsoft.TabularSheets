using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheet.Spreadsheets
{
    public class TabularSpreadsheet<T> : TabularData<T>
    {

        public string Title { get; set; } = "Sheet";
        public HeaderOptions Header { get; set; } = new();


        public TabularSpreadsheet<T> SetSheetTitle(string title)
        {
            throw new NotImplementedException();
        }

        public void Create(string path) { throw new NotImplementedException(); }
        public MemoryStream Create() { throw new NotImplementedException(); }
    }
}
