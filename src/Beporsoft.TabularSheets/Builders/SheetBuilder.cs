using Beporsoft.TabularSheets.Builders.StyleBuilders;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Builders
{
    internal class SheetBuilder<T>
    {
        public SheetBuilder(TabularSpreadsheet<T> table, StylesheetBuilder styleBuilder)
        {
            Table = table;
            StyleBuilder = styleBuilder;
        }

        public TabularSpreadsheet<T> Table { get; }
        public StylesheetBuilder StyleBuilder { get; }

        public SheetData BuildSheetData()
        {
            SheetData sheetData = new SheetData();

            return sheetData;
        }

        private Row CreateHeader()
        {
            Row header = new Row();
            foreach (var col in Table.Columns)
            {
                
            }
            return header;
        }

        protected static Cell CreateCell(string value, CellValues dataType)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType)
            };
        }

        protected static Cell CreateCell(object value)
        {
            Type type = value.GetType();
            Cell cell = new Cell();
            
            if (type == typeof(string))
            {
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue(Convert.ToString(value) ?? string.Empty);
            }
            else if (type == typeof(int))
            {
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(Convert.ToInt32(value));
            }
            else if (type == typeof(double))
            {
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(Convert.ToDouble(value));
            }
            else if (type == typeof(float))
            {
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(Convert.ToDecimal(value));
            }
            else if (type == typeof(bool))
            {
                cell.DataType = CellValues.Boolean;
                cell.CellValue = new CellValue(Convert.ToBoolean(value));
            }
            else if (type == typeof(DateTime))
            {
                cell.DataType = CellValues.Number;
                cell.StyleIndex = 1;
                cell.CellValue = new CellValue(Convert.ToDateTime(value).ToOADate().ToString());
            }
            else
            {
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue(Convert.ToString(value) ?? string.Empty);
            }

            return cell;
        }
    }
}
