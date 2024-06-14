using Beporsoft.TabularSheets.Builders.SheetBuilders;
using Beporsoft.TabularSheets.Builders.StyleBuilders.Adapters;
using Beporsoft.TabularSheets.Tools;
using DocumentFormat.OpenXml.Office.CoverPageProps;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Beporsoft.TabularSheets.Builders.Import
{
    /// <summary>
    /// Enclosure of content of a workbook
    /// </summary>
    internal class SheetImporter
    {

        public SheetImporter(Stream stream, string? sheetName = null)
        {
            Load(stream, sheetName);
        }

        public SheetImporter(string path, string? sheetName = null)
        {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);
            Load(fs, sheetName);
        }



        public SheetData Data { get; private set; } = null!;
        public Stylesheet Stylesheet { get; private set; } = null!;
        public SharedStringTable SharedStrings { get; private set; } = null!;
        public SheetDimension Dimensions { get; private set; } = null!;
        public AutoFilter? AutoFilter { get; private set; }
        public string Title { get; private set; } = null!;

        public string GetDimensionReference()
        {
            return Dimensions?.Reference?.Value ?? string.Empty;
        }
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

        public List<Cell> GetHeaderCells()
        {
            var cells = Data.Descendants<Cell>()
                .Where(c =>
                {
                    var cellRef = CellRefBuilder.GetIndexes(c.CellReference!);
                    return cellRef.Row == 0;
                });
            return cells.ToList();
        }

        public List<Cell> GetBodyCells()
        {
            var cells = Data.Descendants<Cell>()
                .Where(c =>
                {
                    var cellRef = CellRefBuilder.GetIndexes(c.CellReference!);
                    return cellRef.Row != 0;
                });
            return cells.ToList();
        }

        public string? GetSharedString(int indexString)
        {
            var listItems = SharedStrings.Descendants<SharedStringItem>().ToList();
            var item = listItems.Count > indexString ? listItems[indexString] : null;
            return item?.Text?.Text;
        }
        private void Load(Stream stream, string? sheetName = null)
        {
            using var spreadsheet = SpreadsheetDocument.Open(stream, false);
            WorkbookPart workbookPart = spreadsheet.WorkbookPart!;

            Sheet sheet;
            if (sheetName is null)
                sheet = workbookPart.Workbook.Sheets!.Descendants<Sheet>().First();
            else
                sheet = workbookPart.Workbook.Sheets!.Descendants<Sheet>().First(s => s.Name == sheetName);


            WorksheetPart? worksheetPart = workbookPart.GetPartById(sheet.Id?.Value!) as WorksheetPart;

            Worksheet worksheet = worksheetPart?.Worksheet!;

            Data = worksheet.Descendants<SheetData>()!.Single();
            Stylesheet = workbookPart.WorkbookStylesPart!.Stylesheet;
            SharedStrings = workbookPart.SharedStringTablePart!.SharedStringTable;
            Dimensions = worksheet.Descendants<SheetDimension>().Single();
            AutoFilter = worksheet.Descendants<AutoFilter>().SingleOrDefault();


            Title = sheet.Name!.Value!;
        }
    }


}
