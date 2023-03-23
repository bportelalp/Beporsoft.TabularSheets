using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Tools
{
    internal static class OpenXMLHelpers
    {
        internal static HexBinaryValue BuildHexBinaryFromColor(System.Drawing.Color color)
            => new HexBinaryValue($"FF{color.R:X2}{color.G:X2}{color.B:X2}");

        internal static UInt32Value ToUint32Value(int value) => new UInt32Value(Convert.ToUInt32(value));
    }
}
