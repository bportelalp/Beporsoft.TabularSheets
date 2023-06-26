using DocumentFormat.OpenXml;
using System;
using System.Drawing;

namespace Beporsoft.TabularSheets.Tools
{
    /// <summary>
    /// Extension methods to help conversion between native value types and OpenXml reference types
    /// </summary>
    internal static class OpenXmlValueExtensions
    {
        /// <summary>
        /// Convert <see cref="Color"/> to its equivalente <see cref="HexBinaryValue"/> representation
        /// </summary>
        internal static HexBinaryValue ToOpenXmlHexBinary(this Color color)
            => new HexBinaryValue($"FF{color.R:X2}{color.G:X2}{color.B:X2}");

        /// <summary>
        /// Convert <see cref="int"/> to its equivalent <see cref="UInt32Value"/> representation
        /// </summary>
        internal static UInt32Value ToOpenXmlUInt32(this int value)
            => new UInt32Value(Convert.ToUInt32(value));

        /// <summary>
        /// Convert <see cref="int"/> to its equivalent <see cref="UInt32Value"/> representation
        /// </summary>
        internal static UInt32Value? ToOpenXmlUInt32(this int? value)
            => value is null? null: new UInt32Value(Convert.ToUInt32(value));
    }
}
