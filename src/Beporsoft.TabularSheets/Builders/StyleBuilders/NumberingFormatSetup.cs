using Beporsoft.TabularSheets.Builders.Interfaces;
using Beporsoft.TabularSheets.Tools;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Diagnostics;

namespace Beporsoft.TabularSheets.Builders.StyleBuilders
{
    /// <summary>
    /// Builder for the qualified node x:numFmt of 
    /// <see href="https://www.ecma-international.org/publications-and-standards/standards/ecma-376/">ECMA-376-1:2016 §18.8.30</see>
    /// </summary>
    [DebuggerDisplay("Id={Index} | Pattern={Pattern}")]
    internal class NumberingFormatSetup : Setup, IEquatable<NumberingFormatSetup?>, IIndexedSetup
    {

        public NumberingFormatSetup(string pattern)
        {
            Pattern = pattern;
        }

        public static int StartIndexNotBuiltin = 164;
        public string Pattern { get; set; } = null!;

        public override OpenXmlElement Build()
        {
            var numberingFormat = new NumberingFormat()
            {
                NumberFormatId = Index.ToOpenXmlUInt32(),
                FormatCode = Pattern
            };
            return numberingFormat;
        }

        public static NumberingFormatSetup FromOpenXmlNumberingFormat(NumberingFormat numFormatXml)
        {
            string pattern = numFormatXml.FormatCode!;
            return new NumberingFormatSetup(pattern);
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
