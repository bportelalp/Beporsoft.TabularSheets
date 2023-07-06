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
        private static readonly Dictionary<Type, TypeConverter> _knownConverters = new()
        {
            // Convertible to numeric, all can use CellValues.Number
            {typeof(short), new TypeConverter(CellValues.Number, (obj) => new CellValue(Convert.ToInt32(obj))) },
            {typeof(ushort), new TypeConverter(CellValues.Number, (obj) => new CellValue(Convert.ToInt32(obj))) },
            {typeof(int), new TypeConverter(CellValues.Number, (obj) => new CellValue(Convert.ToInt32(obj))) },
            {typeof(uint), new TypeConverter(CellValues.Number, (obj) => new CellValue(Convert.ToDouble(obj))) },
            {typeof(double), new TypeConverter(CellValues.Number, (obj) => new CellValue(Convert.ToDouble(obj))) },
            {typeof(float), new TypeConverter(CellValues.Number, (obj) => new CellValue(Convert.ToDouble(obj))) },
            {typeof(decimal), new TypeConverter(CellValues.Number, (obj) => new CellValue(Convert.ToDecimal(obj))) },
            {typeof(bool), new TypeConverter(CellValues.Boolean, (obj) => new CellValue(Convert.ToBoolean(obj))) },
            // DateTime must be saved as number, using the value as OADate, It is responsability of formatting to convert it as datetime
            {typeof(DateTime), new TypeConverter(CellValues.Number, (obj) => new CellValue(Convert.ToDateTime(obj).ToOADate())) },
            // TimeSpan must be saved as number using the total days
            {typeof(TimeSpan), new TypeConverter(CellValues.Number, (obj) => new CellValue(((TimeSpan)obj).TotalDays)) }
        };

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
            TypeConverter? converter = null;
            if (_knownConverters.ContainsKey(type))
                converter = _knownConverters[type];

            if (converter is null)
            {
                // Converter null, use a shared string registering it.
                cell.DataType = CellValues.SharedString;
                string valueString = Convert.ToString(value)?? string.Empty;
                int indexSharedTable = SharedStrings.RegisterString(valueString);
                cell.CellValue = new CellValue(indexSharedTable);
            }
            else
            {
                cell.DataType = converter.Type;
                cell.CellValue = converter.Converter.Invoke(value);
            }
            return cell;
        }

        /// <summary>
        /// A class to store the matching of <see cref="CellValues"/> and a function which convert an <see cref="System.Object"/>
        /// to the compatible <see cref="System.Type"/> with it.
        /// </summary>
        private class TypeConverter
        {
            public TypeConverter(CellValues type, Func<object, CellValue> converter)
            {
                Type = type;
                Converter = converter;
            }

            public CellValues Type { get; set; }
            public Func<object, CellValue> Converter { get; set; }
        }
    }
}
