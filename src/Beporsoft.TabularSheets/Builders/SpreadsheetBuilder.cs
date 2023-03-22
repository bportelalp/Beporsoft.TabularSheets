using Beporsoft.TabularSheets.Builders.StyleBuilders;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Builder
{
    internal class SpreadsheetBuilder
    {
        public SpreadsheetBuilder()
        {
            StylesheetBuilder = new StylesheetBuilder();
        }
        public SpreadsheetBuilder(StylesheetBuilder stylesheetBuilder)
        {

            StylesheetBuilder = stylesheetBuilder;

        }

        public StylesheetBuilder StylesheetBuilder { get; }


        public MemoryStream Create<T>(TabularSpreadsheet<T> table)
        {
            MemoryStream stream = new MemoryStream();
            using var spreadsheet = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            return stream;
        }

        public void AddWorkbookPartFromTable<T>(ref WorkbookPart workbookPart, TabularSpreadsheet<T> table)
        {
            
        }

        public SheetData BuildSheetData<T>(TabularSpreadsheet<T> table)
        {

        }

        private Row CreateHeader()
    }
}
