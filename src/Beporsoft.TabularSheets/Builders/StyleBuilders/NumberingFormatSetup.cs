using Beporsoft.TabularSheets.Builders.Interfaces;
using Beporsoft.TabularSheets.Tools;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Diagnostics;

namespace Beporsoft.TabularSheets.Builders.StyleBuilders
{
    [DebuggerDisplay("Id={Index} | Pattern={Pattern}")]
    internal class NumberingFormatSetup : Setup, IEquatable<NumberingFormatSetup?>, IIndexedSetup
    {
        public static int StartIndexNotBuiltin = 164;

        public NumberingFormatSetup(string pattern)
        {
            Pattern = pattern;
        }

        public string Pattern { get; set; } = null!;

        public override OpenXmlElement Build()
        {
            var numberingFormat = new NumberingFormat()
            {
                NumberFormatId = Index.ToUint32OpenXml(),
                FormatCode = Pattern
            };
            return numberingFormat;
        }

        public override bool Equals(object? obj)
        {
            return Equals(obj as NumberingFormatSetup);
        }

        public bool Equals(NumberingFormatSetup? other)
        {
            return other is not null &&
                   Pattern == other.Pattern;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(Pattern);
        }
    }
}
