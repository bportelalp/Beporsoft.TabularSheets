using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Style
{
    public class HeaderStyle
    {
        public Color? BackgroundColor { get; set; } = null;
        public Color FontColor { get; set; } = Color.Black;
        public int FontSize { get; set; } = 10;
        public string FontFamily { get; set; } = null!;

    }
}
