using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Builders
{
    internal struct BuildHelpers
    {
        internal const int DefaultColumnWidth = 10;

        internal static readonly HashSet<Type> NumericTypes = new HashSet<Type>
        {
            typeof(int),  typeof(double),  typeof(decimal),
            typeof(long), typeof(short),   typeof(sbyte),
            typeof(byte), typeof(ulong),   typeof(ushort),
            typeof(uint), typeof(float)
        };

        internal static readonly HashSet<Type> DateTimeTypes = new HashSet<Type>() { typeof(DateTime), typeof(DateTimeOffset) };

        internal static readonly HashSet<Type> TimeSpanTypes = new HashSet<Type>()
        {
            typeof(TimeSpan),
#if NET6_0_OR_GREATER 
            typeof(TimeOnly) 
#endif 
        };

        internal static readonly HashSet<Type> CellTypesWithStyle = new HashSet<Type>()
        {
            typeof(DateTime),
            typeof(DateTimeOffset),
            typeof(TimeSpan),
#if NET6_0_OR_GREATER 
            typeof(TimeOnly) 
#endif 
        };
    }
}
