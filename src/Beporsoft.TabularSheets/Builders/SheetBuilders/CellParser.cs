using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;

namespace Beporsoft.TabularSheets.Builders.SheetBuilders
{
    internal class CellParser
    {

        public CellParser(SharedStringTable sharedStrings)
        {
            SharedStrings = sharedStrings;
        }

        public SharedStringTable SharedStrings { get; }

        public object? GetValue(Cell cell, Type targetType)
        {
            if(cell.CellValue is null)
                return targetType.IsValueType ? Activator.CreateInstance(targetType) : null;
            
            if(CellTypeConverter.KnownConverters.TryGetValue(targetType, out CellTypeConverter? converter))
            {
                if(converter.Type == cell.DataType!.Value)
                {
                    return converter.ReverseConverter.Invoke(cell.CellValue, SharedStrings);
                }
                else
                {
                    //TODO
                    return SharedStrings.ElementAt(int.Parse(cell.CellValue.InnerText)).InnerText;
                }
            }
            else
            {
                return SharedStrings.ElementAt(int.Parse(cell.CellValue.InnerText)).InnerText;
            }
        }
    }
}
