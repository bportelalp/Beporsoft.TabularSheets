using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Builders
{
    internal struct BuildHelpers
    {
        /// <summary>
        /// Default column width of a standard column, in terms of number of digit characters
        /// </summary>
        public const int DefaultColumnWidth = 10;

        /// <summary>
        /// Default font size in points applied for excell by default
        /// </summary>
        public const double DefaultFontSize = 11.0;

        /// <summary>
        /// Collection of <see cref="Type"/> which are numeric types
        /// </summary>
        public static readonly HashSet<Type> NumericTypes = new HashSet<Type>
        {
            typeof(int),  typeof(double),  typeof(decimal),
            typeof(long), typeof(short),   typeof(sbyte),
            typeof(byte), typeof(ulong),   typeof(ushort),
            typeof(uint), typeof(float)
        };

        /// <summary>
        /// Collection of <see cref="Type"/> which are datetime types
        /// </summary>
        public static readonly HashSet<Type> DateTimeTypes = new HashSet<Type>() { typeof(DateTime), typeof(DateTimeOffset) };

        /// <summary>
        /// Collection of <see cref="Type"/> which time span types
        /// </summary>
        public static readonly HashSet<Type> TimeSpanTypes = new HashSet<Type>()
        {
            typeof(TimeSpan),
#if NET6_0_OR_GREATER 
            typeof(TimeOnly) 
#endif 
        };
    }
}
