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

        public CellStyling.Style? GetCellStyle(int indexFormat)
        {
            List<CellFormat> listFormats = Stylesheet.CellFormats!.Descendants<CellFormat>().ToList();
            CellFormat? cellFormat = listFormats.Count > indexFormat ? listFormats[indexFormat] : null;
            if (cellFormat is not null)
            {
                StyleWrapper style = new() { CellFormat = cellFormat };
                if (cellFormat.FillId is not null)
                    style.Fill = GetFill(Convert.ToInt32(cellFormat.FillId.Value));
                if (cellFormat.FontId is not null)
                    style.Font = GetFont(Convert.ToInt32(cellFormat.FontId.Value));
                if (cellFormat.BorderId is not null)
                    style.Border = GetBorder(Convert.ToInt32(cellFormat.BorderId.Value));
                if (cellFormat.NumberFormatId is not null)
                    style.NumberingFormat = GetNumberingFormat(Convert.ToInt32(cellFormat.NumberFormatId.Value));
                if (cellFormat.Alignment is not null)
                    style.Alignment = cellFormat.Alignment;
                return style.ToStyle();
            }
            else
            {
                return null;
            }
        }

        private Fill GetFill(int indexFill)
        {
            List<Fill> listFills = Stylesheet.Fills!.Descendants<Fill>().ToList();
            return listFills[indexFill];
        }

        private Font GetFont(int indexFont)
        {
            List<Font> listFonts = Stylesheet.Fonts!.Descendants<Font>().ToList();
            return listFonts[indexFont];

        }

        private Border GetBorder(int indexBorder)
        {
            List<Border> listFills = Stylesheet.Borders!.Descendants<Border>().ToList();
            return listFills[indexBorder];

        }

        private NumberingFormat GetNumberingFormat(int indexNumbering)
        {
            NumberingFormat? numberingPattern = Stylesheet.NumberingFormats!
                .Descendants<NumberingFormat>()
                .SingleOrDefault(nf => nf.NumberFormatId!.Value == indexNumbering.ToOpenXmlUInt32());
            if (numberingPattern is null)
            {
                if (NumberingFormatSetupCollection.PredefinedFormats.ContainsKey(indexNumbering))
                {
                    numberingPattern = new()
                    {
                        NumberFormatId = indexNumbering.ToOpenXmlUInt32(),
                        FormatCode = NumberingFormatSetupCollection.PredefinedFormats[indexNumbering]
                    };
                }
            }
            return numberingPattern!;
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


    internal class StyleWrapper
    {
        public CellFormat CellFormat { get; set; } = null!;
        public Border? Border { get; set; }
        public Fill? Fill { get; set; }
        public Font? Font { get; set; }
        public NumberingFormat? NumberingFormat { get; set; }
        public Alignment? Alignment { get; set; }

        public CellStyling.Style ToStyle()
        {
            CellStyling.Style style = new();
            if (Border is not null)
                style.Border = StyleBuilders.BorderSetup.FromOpenXml(Border).BorderStyle;

            if (Fill is not null)
                style.Fill = StyleBuilders.FillSetup.FromOpenXml(Fill).Fill;

            if (Font is not null)
                style.Font = StyleBuilders.FontSetup.FromOpenXml(Font).FontStyle;

            if (NumberingFormat is not null)
                style.NumberingPattern = StyleBuilders.NumberingFormatSetup.FromOpenXml(NumberingFormat).Pattern;

            if (Alignment is not null)
                style.Alignment = StyleBuilders.FormatSetup.FromOpenXml(Alignment);

            return style;
        }
    }
}
