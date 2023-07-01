using Beporsoft.TabularSheets.Builders.SheetBuilders;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Beporsoft.TabularSheets.Tools;

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
                if(cellFormat.NumberFormatId is not null)
                    style.NumberingFormat = GetNumberingFormat(Convert.ToInt32(cellFormat.NumberFormatId.Value));
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
            NumberingFormat numberingPattern = Stylesheet.NumberingFormats!
                .Descendants<NumberingFormat>()
                .Single(nf => nf.NumberFormatId!.Value == indexNumbering.ToOpenXmlUInt32());
            return numberingPattern;
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


    internal class StyleWrapper
    {
        public CellFormat CellFormat { get; set; } = null!;
        public Border? Border { get; set; }
        public Fill? Fill { get; set; }
        public Font? Font { get; set; }
        public NumberingFormat? NumberingFormat { get; set; }

        public Beporsoft.TabularSheets.CellStyling.Style ToStyle()
        {
            CellStyling.Style style = new();
            if (Border is not null)
            {
                // Left Border
                if (Border.LeftBorder is not null)
                {

                }
            }
            if (Fill is not null)
            {
                style.Fill.BackgroundColor = Fill.PatternFill?.ForegroundColor?.Rgb?.FromOpenXmlHexBinaryValue();
            }
            if (Font is not null)
            {
                style.Font.Font = Font.FontName?.Val?.Value;
                style.Font.Color = Font.Color?.Rgb?.FromOpenXmlHexBinaryValue();
                style.Font.Size = Font.FontSize?.Val?.Value is not null ? Convert.ToInt32(Font.FontSize?.Val?.Value) : null;
            }
            if(NumberingFormat is not null)
            {
                style.NumberingPattern = NumberingFormat.FormatCode;
            }

            return style;
        }
    }
}
