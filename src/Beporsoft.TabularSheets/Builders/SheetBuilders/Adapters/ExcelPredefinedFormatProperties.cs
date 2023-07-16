using DocumentFormat.OpenXml.Spreadsheet;

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
