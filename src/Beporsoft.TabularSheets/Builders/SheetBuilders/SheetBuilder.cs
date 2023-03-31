using Beporsoft.TabularSheets.Builders.StyleBuilders;
using Beporsoft.TabularSheets.CellStyling;
using Beporsoft.TabularSheets.Tools;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;

namespace Beporsoft.TabularSheets.Builders.SheetBuilders
{
    /// <summary>
    /// The class which handles the creation of the <see cref="SheetData"/> which represent the <see cref="Table"/> and register
    /// the shared elements like shared strings and styles inside the containers <see cref="SharedStrings"/> and <see cref="StyleBuilder"/>
    /// respectively.
    /// </summary>
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

        #region Properties

        /// <summary>
        /// The <see cref="TabularSheet{T}"/> used to build the <see cref="SheetData"/>
        /// </summary>
        public TabularSheet<T> Table { get; }

        /// <summary>
        /// The container which stores the generated styles
        /// </summary>
        public StylesheetBuilder StyleBuilder { get; }

        /// <summary>
        /// The container which stores the shared strings
        /// </summary>
        public SharedStringBuilder SharedStrings { get; }

        /// <summary>
        /// The cell builder, which helps on building <see cref="Cell"/> instances
        /// </summary>
        public CellBuilder CellBuilder { get; }
        #endregion

        #region Public
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
        #endregion

        #region Create SheetData components
        /// <summary>
        /// Create the <see cref="Row"/> which represents the header of the table, including the registry of
        /// styles inside <see cref="StyleBuilder"/>
        /// </summary>
        private Row CreateHeaderRow()
        {
            Row header = new();

            header.RowIndex = OpenXmlHelpers.ToUint32Value(_cellRefIterator.CurrentRow + 1);


            int? formatId = RegisterHeaderStyle();

            _cellRefIterator.ResetCol();
            foreach (var col in Table.Columns)
            {
                Cell cell = CellBuilder.CreateCell(col.Title);
                if (formatId is not null)
                    cell.StyleIndex = OpenXmlHelpers.ToUint32Value(formatId.Value);

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
            row.RowIndex = OpenXmlHelpers.ToUint32Value(_cellRefIterator.CurrentRow + 1);
            _cellRefIterator.ResetCol();
            foreach (var col in Table.Columns)
            {
                object value = col.Apply(item);
                Cell cell = new();
                if (value is not null)
                {
                    cell = BuildDataCell(value);
                }
                cell.CellReference = _cellRefIterator.MoveNextColAfter();
                row.Append(cell);
            }
            return row;
        }


        private Cell BuildDataCell(object value)
        {
            Cell cell = CellBuilder.CreateCell(value);
            FillSetup? fillSetup = null;
            FontSetup? fontSetup = null;
            BorderSetup? borderSetup = null;
            NumberingFormatSetup? numberingFormatSetup = null;

            if (!Table.DefaultStyle.Fill.Equals(FillStyle.Default))
                fillSetup = new FillSetup(Table.DefaultStyle.Fill);

            if (!Table.DefaultStyle.Font.Equals(FontStyle.Default))
                fontSetup = new FontSetup(Table.DefaultStyle.Font);

            if (!Table.DefaultStyle.Border.Equals(BorderStyle.Default))
                borderSetup = new BorderSetup(Table.DefaultStyle.Border);

            if (value.GetType() == typeof(DateTime))
            {
                numberingFormatSetup = new NumberingFormatSetup(Table.DefaultStyle.DateTimeFormat);
            }
            int formatId = StyleBuilder.RegisterFormat(fillSetup, fontSetup, borderSetup, numberingFormatSetup);
            cell.StyleIndex = OpenXmlHelpers.ToUint32Value(formatId);
            return cell;
        }

        private int? RegisterHeaderStyle()
        {
            FillSetup? fill = null;
            FontSetup? font = null;
            BorderSetup? border = null;

            FillStyle combinedFill = StyleCombiner.Combine(Table.HeaderStyle.Fill, Table.DefaultStyle.Fill);
            if (!combinedFill.Equals(FillStyle.Default))
                fill = new FillSetup(combinedFill);

            FontStyle combinedFont = StyleCombiner.Combine(Table.HeaderStyle.Font, Table.DefaultStyle.Font);
            if (!combinedFill.Equals(FontStyle.Default))
                font = new FontSetup(combinedFont);

            BorderStyle combinedBorder = StyleCombiner.Combine(Table.HeaderStyle.Border, Table.DefaultStyle.Border);
            if (!combinedBorder.Equals(BorderStyle.Default))
                border = new BorderSetup(combinedBorder);

            if (font is null && fill is null)
                return null;

            int? formatId = StyleBuilder.RegisterFormat(fill, font, border);

            return formatId;
        }
        #endregion
    }
}
