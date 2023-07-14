using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Builders.SheetBuilders.Adapters
{

    internal class ExcelPredefinedFormatProperties
    {
        public static SheetFormatProperties Create()
        {
            return new SheetFormatProperties()
            {
                BaseColumnWidth = 10,
                DefaultRowHeight = 15
            };
        }
    }
}
