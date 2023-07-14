using Beporsoft.TabularSheets.Builders.Interfaces;
using Beporsoft.TabularSheets.CellStyling;
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
        public FormatSetup(FillSetup? fill, FontSetup? font, BorderSetup? border, NumberingFormatSetup? numberingFormat, AlignmentStyle? alignment)
        {
            Fill = fill;
            Font = font;
            Border = border;
            NumberingFormat = numberingFormat;
            Alignment = alignment;
        }

        public FillSetup? Fill { get; set; }
        public FontSetup? Font { get; set; }
        public BorderSetup? Border { get; set; }
        public NumberingFormatSetup? NumberingFormat { get; set; }
        public AlignmentStyle? Alignment { get; set; }

        public override OpenXmlElement Build()
        {
            var format = new CellFormat
            {
                FillId = GetSetupId(Fill),
                FontId = GetSetupId(Font),
                BorderId = GetSetupId(Border),
                NumberFormatId = GetSetupId(NumberingFormat),
                Alignment = BuildAlignment()


            };

            format.ApplyFill = format.FillId is not null;
            format.ApplyFont = format.FontId is not null;
            format.ApplyBorder = format.BorderId is not null;
            format.ApplyNumberFormat = format.NumberFormatId is not null;
            format.ApplyAlignment = format.Alignment is not null;

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
                   EqualityComparer<NumberingFormatSetup?>.Default.Equals(NumberingFormat, other.NumberingFormat) &&
                   EqualityComparer<AlignmentStyle?>.Default.Equals(Alignment, other.Alignment);
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(Fill, Font, Border, Alignment);
        }

        #region Create OpenXmlElement
        private static UInt32Value? GetSetupId(Setup? setup)
        {
            UInt32Value? value = null;
            if (setup is not null)
                value = setup.Index.ToOpenXmlUInt32();
            return value;
        }

        private Alignment? BuildAlignment()
        {
            Alignment? align = null;
            if (Alignment is not null)
            {
                bool horizontalOk = Enum.TryParse(Alignment.Horizontal.ToString(), out HorizontalAlignmentValues horizontal);
                bool verticalOk = Enum.TryParse(Alignment.Vertical.ToString(), out VerticalAlignmentValues vertical);
                align = new Alignment()
                {
                    Horizontal = horizontalOk ? horizontal : null,
                    Vertical = verticalOk ? vertical : null,
                    WrapText = Alignment.TextWrap,
                };
            }
            return align;
        }
        #endregion
    }
}
