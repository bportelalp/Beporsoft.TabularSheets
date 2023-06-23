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
        /// Create an instance of <typeparamref name="T"/> filling its properties from <paramref name="highestPriority"/>, unless they're
        /// <see langword="null"/>. In this case, use the equivalent property from <paramref name="lowestPriority"/>
        /// </summary>
        internal static T Combine<T>(T highestPriority, T lowestPriority) where T : class, new()
        {
            PropertyInfo[] properties = typeof(T).GetProperties();
            T result = new T();
            foreach (PropertyInfo property in properties)
            {
                Type propertyType = property.PropertyType;
                if (propertyType.IsValueType || propertyType == typeof(string))
                {
                    object? assignedValue;
                    object? priorityValue = property.GetValue(highestPriority);
                    if (priorityValue is not null)
                        assignedValue = priorityValue;
                    else
                        assignedValue = property.GetValue(lowestPriority);

                    property.SetValue(result, assignedValue);
                }
                else
                {
                    MethodInfo? methodCombineInfo = typeof(StyleCombiner).GetMethod(nameof(Combine), BindingFlags.Static | BindingFlags.NonPublic);
                    MethodInfo method = methodCombineInfo!.MakeGenericMethod(propertyType);
                    object? propertyInstanceHighest = property.GetValue(highestPriority);
                    object? propertyInstanceLowest = property.GetValue(lowestPriority);
                    var resultInstance = method.Invoke(null, new object?[] { propertyInstanceHighest, propertyInstanceLowest });
                    property.SetValue(result, resultInstance);
                }
            }
            return result;
        }
    }
}
