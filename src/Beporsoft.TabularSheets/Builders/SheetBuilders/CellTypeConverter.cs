using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
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
        private CellTypeConverter(CellValues type, 
            Func<object, SharedStringBuilder, CellValue> forwardConverter,
            Func<CellValue, SharedStringTable?, object> reverseConverter)
        {
            Type = type;
            ForwardConverter = forwardConverter;
            ReverseConverter = reverseConverter;
        }

        /// <summary>
        /// The OpenXml datatype which applies this converter,
        /// </summary>
        public CellValues Type { get; set; }

        /// <summary>
        /// Forward function to convert from dotnet datatype to <see cref="CellValue"/>. The function
        /// receives as argument a <see cref="SharedStringBuilder"/> to register shared strings if required.
        /// </summary>
        public Func<object, SharedStringBuilder, CellValue> ForwardConverter { get; set; }

        /// <summary>
        /// Reverse function to convert from <see cref="CellValue"/> to dotnet type. The function
        /// receives as argument a <see cref="SharedStringTable"/> to lookup for shared strings if required.
        /// </summary>
        public Func<CellValue, SharedStringTable?, object> ReverseConverter { get; set; }


        /// <summary>
        /// All Well-known converters between dotnet primitive types and <see cref="CellValues"/>.
        /// </summary>
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
                (value, t) => (double)Convert.ToDouble(value.InnerText, CultureInfo.InvariantCulture))},

            {typeof(float), new CellTypeConverter(CellValues.Number,
                (obj, tb) => new CellValue(Convert.ToDouble(obj)),
                (value, t) => (float)Convert.ToDouble(value.InnerText, CultureInfo.InvariantCulture))},

            {typeof(decimal), new CellTypeConverter(CellValues.Number,
                (obj, tb) => new CellValue(Convert.ToDecimal(obj)),
                (value, t) => (decimal)Convert.ToDecimal(value.InnerText, CultureInfo.InvariantCulture))},

            {typeof(bool), new CellTypeConverter(CellValues.Boolean,
                (obj, tb) => new CellValue(Convert.ToBoolean(obj)),
                (value, t) => (bool)Convert.ToBoolean(value.InnerText))},

            // DateTime must be saved as number, using the value as OADate, It is responsability of formatting to convert it as datetime
            {typeof(DateTime), new CellTypeConverter(CellValues.Number,
                (obj, tb) => new CellValue(Convert.ToDateTime(obj).ToOADate()),
                (value, t) => DateTime.FromOADate(Convert.ToDouble(value.InnerText, CultureInfo.InvariantCulture)))},

            {typeof(DateTimeOffset), new CellTypeConverter(CellValues.Number,
                (obj, tb) => new CellValue(((DateTimeOffset)obj).DateTime.ToOADate()),
                (value, t) => DateTime.FromOADate(Convert.ToDouble(value.InnerText, CultureInfo.InvariantCulture))) },

            // TimeSpan must be saved as number using the total days
            {typeof(TimeSpan), new CellTypeConverter(CellValues.Number,
                (obj, tb) => new CellValue(((TimeSpan)obj).TotalDays),
                (value, t) =>
                {
                    double doubleValue = Convert.ToDouble(value.InnerText, CultureInfo.InvariantCulture);
                    TimeSpan time = TimeSpan.FromDays(doubleValue);
                    return time;
                })
            },
            
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
                })
            },

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
                })
            },

#if NET6_0_OR_GREATER
            // TODO datos especiales
            {typeof(TimeOnly), new CellTypeConverter(CellValues.Number,
                (obj, t) => new CellValue(((TimeOnly)obj).ToTimeSpan().TotalDays),
                (value, t) =>
                {
                    double doubleValue = Convert.ToDouble(value.InnerText, CultureInfo.InvariantCulture);
                    TimeSpan timeSpan = TimeSpan.FromDays(doubleValue);
                    TimeOnly timeOnly = TimeOnly.FromTimeSpan(timeSpan);
                    return timeOnly;
                })
            },
            {typeof(DateOnly), new CellTypeConverter(CellValues.Number,
                (obj, tb) => new CellValue(Convert.ToDateTime(obj).ToOADate()),
                (value, t) =>
                {
                    double doubleValue = Convert.ToDouble(value.InnerText, CultureInfo.InvariantCulture);
                    DateTime dateTime = DateTime.FromOADate(doubleValue);
                    DateOnly dateOnly = DateOnly.FromDateTime(dateTime);
                    return dateOnly;
                })
            },
#endif
        };
    }
}
