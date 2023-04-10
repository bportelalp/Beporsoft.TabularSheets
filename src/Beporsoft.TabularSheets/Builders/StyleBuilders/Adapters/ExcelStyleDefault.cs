using Beporsoft.TabularSheets.CellStyling;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Builders.StyleBuilders.Adapters
{
    /// <summary>
    /// A class to create the default styles used by Microsoft Excel which is required to correct rendering.
    /// </summary>
    internal class ExcelStyleDefaults
    {
        internal List<FillSetup> Fills { get; set; } = new();
        internal List<FontSetup> Fonts { get; set; } = new();
        internal List<BorderSetup> Border { get; set; } = new();
        internal List<FormatSetup> Formats { get; set; } = new();
        internal List<NumberingFormatSetup> NumberingFormats { get; set; } = new();

        internal static ExcelStyleDefaults Create()
        {
            ExcelStyleDefaults defaults = new ExcelStyleDefaults();
            FillSetup fillNone = new FillSetup(new FillStyle() { PatternValue = DocumentFormat.OpenXml.Spreadsheet.PatternValues.None });
            FillSetup fillGray125 = new FillSetup(new FillStyle() { PatternValue = DocumentFormat.OpenXml.Spreadsheet.PatternValues.Gray125 });
            FontSetup fontCalibri12 = new FontSetup(new FontStyle()
            {
                Font = "Calibri",
                Size = 12
            });
            BorderSetup borderEmpty = new BorderSetup(new BorderStyle());
            NumberingFormatSetup numFormatGeneral = new NumberingFormatSetup("General");
            FormatSetup formatEmpty = new FormatSetup(fillNone, fontCalibri12, borderEmpty, numFormatGeneral);

            defaults.Fills.Add(fillNone);
            defaults.Fills.Add(fillGray125);
            defaults.Fonts.Add(fontCalibri12);
            defaults.Border.Add(borderEmpty);
            defaults.NumberingFormats.Add(numFormatGeneral);
            defaults.Formats.Add(formatEmpty);
            return defaults;
        }
    }
}
