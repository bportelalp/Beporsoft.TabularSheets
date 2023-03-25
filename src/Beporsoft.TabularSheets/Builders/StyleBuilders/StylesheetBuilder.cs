using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Builders.StyleBuilders
{
    internal class StylesheetBuilder
    {
        private readonly StyleSetupCollection<FillSetup> _fills = new();
        private readonly StyleSetupCollection<FontSetup> _fonts = new();
        private readonly StyleSetupCollection<BorderSetup> _borders = new();
        private readonly StyleSetupCollection<FormatSetup> _formats = new();

        public int RegisterFormat(FillSetup fill) => RegisterFormat(fill, null, null);
        public int RegisterFormat(FontSetup font) => RegisterFormat(null, font, null);
        public int RegisterFormat(BorderSetup border) => RegisterFormat(null, null, border);

        /// <summary>
        /// Register a cell format with the parameters provided
        /// </summary>
        /// <returns>
        /// The index which represents this style and must be linked to the cell. Repeated calls which match the <see cref="EqualityComparer{T}"/>
        /// of setups will return the same index
        /// </returns>
        public int RegisterFormat(FillSetup? fill, FontSetup? font, BorderSetup? border)
        {
            if (fill is not null)
                _fills.Register(fill);
            if (font is not null)
                _fonts.Register(font);
            if (border is not null)
                _borders.Register(border);

            var format = new FormatSetup(fill, font, border);
            var formatId = _formats.Register(format);
            return formatId;
        }

        public Fills GetFills() => _fills.GetContainer<Fills>();
        public CellFormats GetFormats() => _formats.GetContainer<CellFormats>();
        public Fonts GetFonts() => _fonts.GetContainer<Fonts>();



    }
}
