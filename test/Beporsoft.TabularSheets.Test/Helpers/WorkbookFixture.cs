using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Test.Helpers
{
    internal class WorkbookFixture
    {
        public WorkbookFixture(Stream stream)
        {
            Load(stream);
        }

        public WorkbookFixture(string path)
        {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);
            Load(fs);
        }

        public Dictionary<string, SheetFixture> Sheets { get; set; } = new();

        private void Load(Stream stream)
        {
            using var spreadsheet = SpreadsheetDocument.Open(stream, false);
            WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
            Sheets.Clear();
            foreach (var sheet in workbookPart.Workbook.Sheets!.Descendants<Sheet>())
            {
                string? name = sheet.Name;
                if (name is not null)
                    Sheets.Add(name, new SheetFixture(stream, name));
            }
        }
    }
}
