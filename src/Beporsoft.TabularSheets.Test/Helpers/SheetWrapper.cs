using Beporsoft.TabularSheets.Builders.SheetBuilders;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Test.Helpers
{
    /// <summary>
    /// Enclosure of content of a workbook
    /// </summary>
    internal class SheetWrapper
    {

        public SheetWrapper(string path)
        {
            Load(path);
        }

        public SheetData Data { get; private set; } = null!;
        public Stylesheet Stylesheet { get; private set; } = null!;
        public SharedStringTable SharedStrings { get; private set; } = null!;

        public Cell GetHeaderCellByColumn(int col)
        {
            string headerCellRef = CellRefBuilder.BuildRef(0, col);
            var cellHeader = Data.Descendants<Cell>()
                .Single(c => c.CellReference == headerCellRef);
            return cellHeader;
        }

        public List<Cell> GetBodyCellsByColumn(int col)
        {
            var cells = Data.Descendants<Cell>()
                .Where(c =>
                {
                    var cellRef = CellRefBuilder.GetIndexes(c.CellReference!);
                    return cellRef.Row != 0 && cellRef.Col == col;
                });

            return cells.ToList();
        }

        public string? GetSharedString(int indexString)
        {
            var listItems = SharedStrings.Descendants<SharedStringItem>().ToList();
            var item = listItems.Count > indexString ? listItems[indexString] : null;
            return item?.Text?.Text;
        }

        private void Load(string filePath)
        {
            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
                Worksheet worksheet = workbookPart.WorksheetParts.First().Worksheet;
                Data = worksheet.Descendants<SheetData>()!.Single();
                Stylesheet = workbookPart.WorkbookStylesPart!.Stylesheet;
                SharedStrings = workbookPart.SharedStringTablePart!.SharedStringTable;
            }
        }
    }
}
