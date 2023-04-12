using Beporsoft.TabularSheets.Builders.Interfaces;
using Beporsoft.TabularSheets.Tools;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace Beporsoft.TabularSheets.Builders.StyleBuilders
{
    /// <summary>
    /// Builder for the qualified node x:xf of 
    /// <see href="https://www.ecma-international.org/publications-and-standards/standards/ecma-376/">ECMA-376-1:2016 §18.8.45</see>
    /// </summary>
    [DebuggerDisplay("Id={Index} | Fill={Fill is not null} | Font={Font is not null} | Border={Border is not null} | Numb={NumberingFormat is not null}")]
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
            if (format.FillId is null)
                format.ApplyFill = false;
            else
                format.ApplyFill = true;

            if (format.FontId is null)
                format.ApplyFont = false;
            else
                format.ApplyFont = true;

            if (format.BorderId is null)
                format.ApplyBorder = false;
            else
                format.ApplyBorder = true;

            if (format.NumberFormatId is null)
                format.ApplyNumberFormat = false;
            else
                format.ApplyNumberFormat = true;

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
                   EqualityComparer<BorderSetup?>.Default.Equals(Border, other.Border) &&
                   EqualityComparer<NumberingFormatSetup?>.Default.Equals(NumberingFormat, other.NumberingFormat);
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
                value = setup.Index.ToUint32OpenXml();
            return value;
        }
        #endregion
    }
}
