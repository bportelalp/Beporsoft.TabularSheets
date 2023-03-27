using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Style
{
    public class TabularSheetOptions
    {
        public string DefaultDateTimeFormat { get; set; } = $"{CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern} hh:mm:ss";
        public FontStyle DefaultFont { get; set; } = new FontStyle();
        public FillStyle DefaultFill { get; set; } = new FillStyle();

    }
}
