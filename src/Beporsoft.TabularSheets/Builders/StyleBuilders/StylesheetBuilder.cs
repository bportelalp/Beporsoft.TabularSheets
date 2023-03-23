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
        private readonly StyleSetupCollection<FormatSetup> _formats = new();


        public int RegisterFormat(System.Drawing.Color color)
        {
            var fill = new FillSetup(color, null);
            _fills.Register(fill);
            var format = new FormatSetup(fill, null);
            var formatId = _formats.Register(format);
            return formatId;
        }

        public Fills GetFills() => _fills.GetContainer<Fills>();
        public CellFormats GetFormats() => _formats.GetContainer<CellFormats>();



    }
}
