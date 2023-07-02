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
        /// Method <see cref="Combine{T}(T, T)"/> obtained by reflection to invoke for combine typed objects at runtime.
        /// </summary>
        internal static MethodInfo? MethodCombine = typeof(StyleCombiner).GetMethod(nameof(Combine), BindingFlags.Static | BindingFlags.NonPublic);


        /// <summary>
        /// Create an instance of <typeparamref name="T"/> filling its properties from <paramref name="highestPriority"/>, unless they're
        /// <see langword="null"/>. In this case, use the equivalent property from <paramref name="lowestPriority"/>
        /// </summary>
        internal static T Combine<T>(T highestPriority, T lowestPriority) where T : class, new()
        {
            T result = new T();

            PropertyInfo[] properties = typeof(T).GetProperties();
            foreach (PropertyInfo property in properties)
            {
                Type propertyType = property.PropertyType;
                // TODO - Find a suitable way to filter primitive reference types and string
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
                    // Recursively invoke this method for sibling reference objects using Reflection
                    object? highPriorityByRefProperty = property.GetValue(highestPriority);
                    object? lowPriorityByRefProperty = property.GetValue(lowestPriority);
                    object?[] paramsCombine = new object?[] { highPriorityByRefProperty, lowPriorityByRefProperty };

                    MethodInfo method = MethodCombine!.MakeGenericMethod(propertyType);
                    object? resultByRefProperty = method.Invoke(null, paramsCombine);
                    property.SetValue(result, resultByRefProperty);
                }
            }
            return result;
        }
    }
}
