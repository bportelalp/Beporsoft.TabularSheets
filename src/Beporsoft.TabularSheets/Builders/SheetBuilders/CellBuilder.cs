using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;

namespace Beporsoft.TabularSheets.Builders.SheetBuilders
{
    /// <summary>
    /// Create a cell based on its value. To fill it, use <see cref="TypeConverter"/> to interpret the dotnet value and set the 
    /// specific <see cref="CellValues"/>. In the case of an string, use the <see cref="SharedStringBuilder"/> to register it
    /// and apply cell content to the index of the SharedString
    /// </summary>
    internal class CellBuilder
    {

        public CellBuilder(SharedStringBuilder sharedStrings)
        {
            SharedStrings = sharedStrings;
        }

        /// <summary>
        /// The <see cref="SharedStringBuilder"/> to fill with strings.
        /// </summary>
        public SharedStringBuilder SharedStrings { get; }

        /// <summary>
        /// Create a cell filling <see cref="CellType.DataType"/> and <see cref="CellType.CellValue"/> based on the
        /// interpreted type from dotnet to OpenXmlTypes
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public Cell CreateCell(object value)
        {
            Type type = value.GetType();
            Cell cell = new Cell();
            if (CellTypeConverter.KnownConverters.TryGetValue(type, out CellTypeConverter? converter))
            {
                cell.DataType = converter.Type;
                cell.CellValue = converter.ForwardConverter.Invoke(value, SharedStrings);
            }
            else
            {
                // Converter null, use a shared string registering it.
                cell.DataType = CellValues.SharedString;
                string valueString = Convert.ToString(value) ?? string.Empty;
                int indexSharedTable = SharedStrings.RegisterString(valueString);
                cell.CellValue = new CellValue(indexSharedTable);
            }
            return cell;
        }
    }
}
