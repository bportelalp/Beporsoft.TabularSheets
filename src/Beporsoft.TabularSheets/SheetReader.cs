using Beporsoft.TabularSheets.Builders.Import;
using Beporsoft.TabularSheets.Builders.SheetBuilders;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets
{
    public class SheetReader<T> where T : class, new()
    {
        public SheetReader(TabularSheet<T> table)
        {
            Table = table;
        }

        public TabularSheet<T> Table { get; }

        public void FromSpreadsheet(string spreadsheetPath)
        {
            SpreadsheetImporter<T> importer = new(Table);
            importer.Import(spreadsheetPath);
        }

        public void FromCsv(string csvPath)
        {
        }
    }
}
