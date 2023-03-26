using Beporsoft.TabularSheets.Builders.Interfaces;
using Beporsoft.TabularSheets.Tools;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;

namespace Beporsoft.TabularSheets.Builders.StyleBuilders
{
    internal class FormatSetup : Setup, IEquatable<FormatSetup?>, IIndexedSetup
    {
        public FormatSetup(FillSetup? fill, FontSetup? font, BorderSetup? border, NumberingFormatSetup? numberingFormat)
        {
            Fill = fill;
            Font = font;
            Border = border;
            NumberingFormat = numberingFormat;
        }

        public FillSetup? Fill { get; set; }
        public FontSetup? Font { get; set; }
        public BorderSetup? Border { get; set; }
        public NumberingFormatSetup? NumberingFormat { get; set; }

        public override OpenXmlElement Build()
        {
            var format = new CellFormat
            {
                FillId = GetSetupId(Fill),
                FontId = GetSetupId(Font),
                BorderId = GetSetupId(Border),
                NumberFormatId = GetSetupId(NumberingFormat)
            };
            return format;
        }

        public override bool Equals(object? obj)
        {
            return Equals(obj as FormatSetup);
        }

        public bool Equals(FormatSetup? other)
        {
            return other is not null &&
                   EqualityComparer<FillSetup?>.Default.Equals(Fill, other.Fill) &&
                   EqualityComparer<FontSetup?>.Default.Equals(Font, other.Font) &&
                   EqualityComparer<BorderSetup?>.Default.Equals(Border, other.Border); ;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(Fill, Font, Border);
        }

        #region Assign Ids
        private static UInt32Value? GetSetupId(Setup? setup)
        {
            UInt32Value? value = null;
            if (setup is not null)
                value = OpenXMLHelpers.ToUint32Value(setup.Index);
            return value;
        }
        #endregion
    }
}
