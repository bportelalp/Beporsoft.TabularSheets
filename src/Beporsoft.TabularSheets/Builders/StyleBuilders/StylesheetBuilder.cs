using Beporsoft.TabularSheets.Builders.Interfaces;
using Beporsoft.TabularSheets.Builders.StyleBuilders.Adapters;
using Beporsoft.TabularSheets.Builders.StyleBuilders.SetupCollections;
using Beporsoft.TabularSheets.CellStyling;
using Beporsoft.TabularSheets.Tools;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using static System.Net.WebRequestMethods;

namespace Beporsoft.TabularSheets.Builders.StyleBuilders
{
    /// <summary>
    /// A class to handle the creation of the qualified node x:styleSheet of
    /// <see href="https://www.ecma-international.org/publications-and-standards/standards/ecma-376/">ECMA-376-1:2016 §18.8.39</see>, 
    /// which is the root element of the Styles part.
    /// <br/><br/>
    /// An instance of this class can be consumed by multiple <see cref="TabularSheet{T}"/> to build a workbook, feature that it will be
    /// supported in the future. Although each sheet can contain its own styles, in the 
    /// <see href="https://www.ecma-international.org/publications-and-standards/standards/ecma-376/">ECMA-376-1:2016</see>
    /// a single Styles part is shared for every sheets. This is the reason and requirement to create a single class which handles all the styles.
    /// </summary>
    internal class StylesheetBuilder
    {
        private readonly ISetupCollection<FillSetup> _fills = new IndexedSetupCollection<FillSetup>();
        private readonly ISetupCollection<FontSetup> _fonts = new IndexedSetupCollection<FontSetup>();
        private readonly ISetupCollection<BorderSetup> _borders = new IndexedSetupCollection<BorderSetup>();
        private readonly ISetupCollection<FormatSetup> _formats = new IndexedSetupCollection<FormatSetup>();
        private readonly ISetupCollection<NumberingFormatSetup> _numFormats = new NumberingFormatSetupCollection();
        public StylesheetBuilder()
        {
            InitializeMsExcelDefaults();
        }
        public int RegisteredFills => _fills.Count;
        public int RegisteredFonts => _fonts.Count;
        public int RegisteredBorders => _borders.Count;
        public int RegisteredFormats => _formats.Count;
        public int RegisteredNumberingFormats => _numFormats.Count;

        public int RegisterFormat(FillSetup fill) => RegisterFormat(fill, null, null);
        public int RegisterFormat(FontSetup font) => RegisterFormat(null, font, null);
        public int RegisterFormat(BorderSetup border) => RegisterFormat(null, null, border);
        public int RegisterFormat(NumberingFormatSetup numberingFormat) => RegisterFormat(null, null, null, numberingFormat);

        /// <summary>
        /// Register a cell format with the parameters provided
        /// </summary>
        /// <returns>
        /// The index which represents this style and must be linked to the cell. Repeated calls which match the <see cref="EqualityComparer{T}"/>
        /// of setups will return the same index
        /// </returns>
        public int RegisterFormat(FillSetup? fill, FontSetup? font, BorderSetup? border, NumberingFormatSetup? numberingFormat = null)
        {
            if (fill is not null)
                _fills.Register(fill);
            if (font is not null)
                _fonts.Register(font);
            if (border is not null)
                _borders.Register(border);
            if (numberingFormat is not null)
                _numFormats.Register(numberingFormat);


            var format = new FormatSetup(fill, font, border, numberingFormat);
            var formatId = _formats.Register(format);
            return formatId;
        }

        public Fills GetFills()
        {
            Fills fills = _fills.BuildContainer<Fills>();
            fills.Count = RegisteredFills.ToOpenXmlUInt32();
            return fills;
        }
        public CellFormats GetFormats()
        {
            CellFormats formats = _formats.BuildContainer<CellFormats>();
            formats.Count = RegisteredFormats.ToOpenXmlUInt32();
            return formats;
        }
        public Fonts GetFonts()
        {
            Fonts Fonts = _fonts.BuildContainer<Fonts>();
            Fonts.Count = RegisteredFonts.ToOpenXmlUInt32();
            return Fonts;
        }
        public Borders GetBorders()
        {
            Borders borders = _borders.BuildContainer<Borders>();
            borders.Count = RegisteredBorders.ToOpenXmlUInt32();
            return borders;
        }
        public NumberingFormats GetNumberingFormats()
        {
            NumberingFormats numFormats = _numFormats.BuildContainer<NumberingFormats>();
            numFormats.Count = RegisteredNumberingFormats.ToOpenXmlUInt32();
            return numFormats;
        }

        private void InitializeMsExcelDefaults()
        {
            ExcelPredefinedStyles defaults = ExcelPredefinedStyles.Create();
            foreach (var setup in defaults.PredefinedFills)
                _fills.Register(setup);
            foreach (var setup in defaults.PredefinedFonts)
                _fonts.Register(setup);
            foreach (var setup in defaults.PredefinedBorders)
                _borders.Register(setup);
            foreach (var setup in defaults.PredefinedFormats)
                _formats.Register(setup);
            foreach (var setup in defaults.PredefinedNumberingFormats)
                _numFormats.Register(setup);
        }


    }
}
