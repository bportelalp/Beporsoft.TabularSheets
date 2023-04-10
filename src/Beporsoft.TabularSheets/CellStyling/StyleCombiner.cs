using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.CellStyling
{
    /// <summary>
    /// Allow to create a instance with the properties filled from the values of two instances, 
    /// where the first one has the priority, unless the property was null, so the second one property
    /// will be used.
    /// </summary>
    internal static class StyleCombiner
    {

        /// <summary>
        /// Create an instance of <see cref="{T}"/> filling its properties from <paramref name="highestPriority"/>, unless they're
        /// <see langword="null"/>. In this case, use the equivalent property from <paramref name="lowestPriority"/>
        /// </summary>
        internal static T Combine<T>(T highestPriority, T lowestPriority) where T : class, new()
        {
            PropertyInfo[] properties = typeof(T).GetProperties();
            T result = new T();
            foreach (PropertyInfo property in properties)
            {
                object? assignedValue;
                object? highestValue = property.GetValue(highestPriority);
                object? lowestValue = property.GetValue(lowestPriority);
                if (highestValue is null)
                    assignedValue = lowestValue;
                else
                    assignedValue = highestValue;
                property.SetValue(result, assignedValue);
            }
            return result;
        }
    }
}
