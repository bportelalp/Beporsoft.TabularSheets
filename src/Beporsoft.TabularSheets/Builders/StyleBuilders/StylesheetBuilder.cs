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
        private readonly IndexableStyleCollection<IndexableFill> _fills = new();
        private readonly IndexableStyleCollection<IndexableFont> _fonts = new();
        private readonly IndexableStyleCollection<IndexableFormat> _formats = new();


        public int RegisterFormat(System.Drawing.Color color)
        {
            var fill = new IndexableFill(color, null);
            _fills.Register(fill);
            var format = new IndexableFormat(fill, null);
            var formatId = _formats.Register(format);
            return formatId;
        }

        public Fills GetFills() => _fills.GetContainer<Fills>();
        public CellFormats GetFormats() => _formats.GetContainer<CellFormats>();



    }
}
