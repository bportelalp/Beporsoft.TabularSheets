using Beporsoft.TabularSheets.Builders.StyleBuilders;
using Beporsoft.TabularSheets.Tools;
using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace Beporsoft.TabularSheets.Builders.SheetBuilders
{
    internal class SheetBuilder<T>
    {
        private readonly CellRefIterator _cellRefIterator = new CellRefIterator();
        public SheetBuilder(TabularSheet<T> table, StylesheetBuilder styleBuilder, SharedStringBuilder sharedStrings)
        {
            Table = table;
            StyleBuilder = styleBuilder;
            SharedStrings = sharedStrings;
            CellBuilder = new CellBuilder(SharedStrings);
        }

        public TabularSheet<T> Table { get; }
        public StylesheetBuilder StyleBuilder { get; }
        public SharedStringBuilder SharedStrings { get; }
        public CellBuilder CellBuilder { get; }


        /// <summary>
        /// Build the <see cref="SheetData"/> node using values from <see cref="Table"/>. <br/>
        /// In addition, handles the required styles and include it inside the <see cref="StyleBuilder"/>
        /// for after treatment
        /// </summary>
        /// <returns>An instance of the OpenXml element <see cref="SheetData"/></returns>
        public SheetData BuildSheetData()
        {
            SheetData sheetData = new();
            _cellRefIterator.Reset();
            Row header = CreateHeaderRow();
            sheetData.AppendChild(header);

            foreach (var item in Table.Items)
            {
                _cellRefIterator.MoveNextRow();
                Row row = CreateItemRow(item);
                sheetData.AppendChild(row);
            }
            return sheetData;
        }

        /// <summary>
        /// Create the <see cref="Row"/> which represents the header of the table, including the registry of
        /// styles inside <see cref="StyleBuilder"/>
        /// </summary>
        private Row CreateHeaderRow()
        {
            Row header = new();
            int? formatId = RegisterHeaderStyle();

            _cellRefIterator.ResetCol();
            foreach (var col in Table.Columns)
            {
                Cell cell = CellBuilder.CreateCell(col.Title);
                if (formatId is not null)
                    cell.StyleIndex = OpenXMLHelpers.ToUint32Value(formatId.Value);

                cell.CellReference = _cellRefIterator.MoveNextColAfter();

                header.Append(cell);
            }
            return header;
        }

        /// <summary>
        /// Create the <see cref="Row"/> which represents one item of <see cref="{T}"/> of the table, including the registry of
        /// styles inside <see cref="StyleBuilder"/>
        /// </summary>
        private Row CreateItemRow(T item)
        {
            Row row = new Row();
            _cellRefIterator.ResetCol();
            foreach (var col in Table.Columns)
            {
                object value = col.Apply(item);
                Cell cell = new();
                if (value is not null)
                {
                    cell = CellBuilder.CreateCell(value);
                    // Apply date style
                    if(value.GetType() == typeof(DateTime))
                    {
                        // TODO - Give specific format
                        string datePattern = "mm-dd-yy";//System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.SortableDateTimePattern;
                        var numFormat = new NumberingFormatSetup(datePattern);
                        int nunFmtId = StyleBuilder.RegisterFormat(numFormat);
                        cell.StyleIndex = OpenXMLHelpers.ToUint32Value(nunFmtId);
                    }
                }

                cell.CellReference = _cellRefIterator.MoveNextColAfter();
                row.Append(cell);
            }
            return row;
        }

        private int? RegisterHeaderStyle()
        {
            int? formatId = null;
            FillSetup? fill = null;
            FontSetup? font = null;
            if (Table.HeaderStyle.BackgroundColor is not null)
                fill = new FillSetup(Table.HeaderStyle.BackgroundColor.Value, null);
            font = new FontSetup(Table.HeaderStyle.FontColor, Table.HeaderStyle.FontSize);

            if (font is null && fill is null)
                return null;

            formatId = StyleBuilder.RegisterFormat(fill, font, null);

            return formatId;
        }
    }
}
