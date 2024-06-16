using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Builders.SheetBuilders
{
    /// <summary>
    /// A class to store the matching of <see cref="CellValues"/> and a function which convert an <see cref="System.Object"/>
    /// to the compatible <see cref="System.Type"/> with it.
    /// </summary>
    internal class CellTypeConverter
    {
        private CellTypeConverter(CellValues type, Func<object, SharedStringBuilder, CellValue> converter, Func<CellValue, SharedStringTable?, object> reverseConverter)
        {
            Type = type;
            Converter = converter;
            ReverseConverter = reverseConverter;
        }

        public CellValues Type { get; set; }
        public Func<object, SharedStringBuilder, CellValue> Converter { get; set; }
        public Func<CellValue, SharedStringTable?, object> ReverseConverter { get; set; }


        public static readonly Dictionary<Type, CellTypeConverter> KnownConverters = new()
        { 
            // Convertible to numeric, all can use CellValues.Number
            {typeof(short), new CellTypeConverter(CellValues.Number,
                (obj, tb) => new CellValue(Convert.ToInt32(obj)),
                (value, t) => (short)Convert.ToInt16(value.InnerText))},
            {typeof(ushort), new CellTypeConverter(CellValues.Number,
                (obj, tb) => new CellValue(Convert.ToInt32(obj)),
                (value, t) => (ushort)Convert.ToUInt16(value.InnerText))},
            {typeof(int), new CellTypeConverter(CellValues.Number,
                (obj, tb) => new CellValue(Convert.ToInt32(obj)),
                (value, t) => (int)Convert.ToInt32(value.InnerText))},
            {typeof(uint), new CellTypeConverter(CellValues.Number,
                (obj, tb) => new CellValue(Convert.ToDouble(obj)),
                (value, t) => (uint)Convert.ToUInt16(value.InnerText))},
            {typeof(double), new CellTypeConverter(CellValues.Number,
                (obj, tb) => new CellValue(Convert.ToDouble(obj)),
                (value, t) => (double)Convert.ToDouble(value.InnerText))},
            {typeof(float), new CellTypeConverter(CellValues.Number,
                (obj, tb) => new CellValue(Convert.ToDouble(obj)),
                (value, t) => (float)Convert.ToDouble(value.InnerText))},
            {typeof(decimal), new CellTypeConverter(CellValues.Number,
                (obj, tb) => new CellValue(Convert.ToDecimal(obj)),
                (value, t) => (decimal)Convert.ToDecimal(value.InnerText))},
            {typeof(bool), new CellTypeConverter(CellValues.Boolean,
                (obj, tb) => new CellValue(Convert.ToBoolean(obj)),
                (value, t) => (bool)Convert.ToBoolean(value.InnerText))},

            // DateTime must be saved as number, using the value as OADate, It is responsability of formatting to convert it as datetime
            {typeof(DateTime), new CellTypeConverter(CellValues.Number,
                (obj, tb) => new CellValue(Convert.ToDateTime(obj).ToOADate()),
                (value, t) => DateTime.FromOADate(Convert.ToDouble(value.InnerText)))},
            {typeof(DateTimeOffset), new CellTypeConverter(CellValues.Number,
                (obj, tb) => new CellValue(((DateTimeOffset)obj).DateTime.ToOADate()),
                (value, t) => DateTime.FromOADate(Convert.ToDouble(value.InnerText))) },

            // TimeSpan must be saved as number using the total days
            {typeof(TimeSpan), new CellTypeConverter(CellValues.Number,
                (obj, tb) => new CellValue(((TimeSpan)obj).TotalDays),
                (value, t) => TimeSpan.FromDays(Convert.ToDouble(value.InnerText))) },
#if NET6_0_OR_GREATER
            // TODO datos especiales
            //{typeof(TimeOnly), new CellTypeConverter(CellValues.Number, (obj) => new CellValue(((TimeOnly)obj).ToTimeSpan().TotalDays)) },
            //{typeof(DateOnly), new CellTypeConverter(CellValues.Number, (obj) => new CellValue(Convert.ToDateTime(obj).ToOADate()))}
#endif

            // Shared string data.
            {typeof(string), new CellTypeConverter(CellValues.SharedString,
                (obj, tb) =>
                {
                    string valueString = Convert.ToString(obj) ?? string.Empty;
                    int indexSharedTable = tb.RegisterString(valueString);
                    return new CellValue(indexSharedTable);
                },
                (value, t) => {
                    int index = int.Parse(value.InnerText);
                    string text = t!.ElementAt(index).InnerText;
                    return Convert.ToString(text);
                })},
            {typeof(Guid), new CellTypeConverter(CellValues.SharedString,
                (obj, tb) =>
                {
                    string valueString = Convert.ToString(obj) ?? string.Empty;
                    int indexSharedTable = tb.RegisterString(valueString);
                    return new CellValue(indexSharedTable);
                },
                (value, t) => {
                    int index = int.Parse(value.InnerText);
                    string text = t!.ElementAt(index).InnerText;
                    return new Guid(text);
                })},
        };
    }
}
