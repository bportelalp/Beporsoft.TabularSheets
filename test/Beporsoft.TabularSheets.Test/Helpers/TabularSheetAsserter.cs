using Beporsoft.TabularSheets.Builders;
using Beporsoft.TabularSheets.Builders.SheetBuilders;
using Beporsoft.TabularSheets.Test.TestModels;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Test.Helpers
{
    internal static class TabularSheetAsserter
    {

        /// <summary>
        /// Verify the data on every column is OK as expected
        /// </summary>
        /// <param name="table"></param>
        /// <param name="column"></param>
        /// <param name="sheet"></param>
        internal static void AssertTabularSheet<T>(TabularSheet<T> table, SheetFixture sheet)
        {
            Assert.That(sheet.Title, Is.EqualTo(table.Title));
            foreach (var col in table.Columns)
            {
                AssertColumnHeaderData(col, sheet);
                AssertColumnBodyData(table, col, sheet);
            }
            AssertDimensions(table, sheet);
        }

        internal static void AssertColumnHeaderData<T>(TabularDataColumn<T> column, SheetFixture sheet)
        {
            DocumentFormat.OpenXml.Spreadsheet.Cell headerCell = sheet.GetHeaderCellByColumn(column.Index);
            Assert.Multiple(() =>
            {
                Assert.That(headerCell.InnerText, Is.Not.Null);
                Assert.That(headerCell.DataType!.Value, Is.EqualTo(DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString));
                var indexSharedString = Convert.ToInt32(headerCell.InnerText);
                string? headerTitle = sheet.GetSharedString(indexSharedString);
                Assert.That(headerTitle, Is.EqualTo(column.Title));
            });
        }

        internal static void AssertColumnBodyData<T>(TabularSheet<T> table, TabularDataColumn<T> column, SheetFixture sheet)
        {
            List<DocumentFormat.OpenXml.Spreadsheet.Cell> bodyCells = sheet.GetBodyCellsByColumn(column.Index);
            foreach (var cell in bodyCells)
            {
                Assert.Multiple(() =>
                {
                    int row = CellRefBuilder.GetRowIndex(cell.CellReference!);
                    T item = table[row - 1];
                    object value = column.Apply(item);
                    if (cell.DataType!.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString)
                    {
                        var indexSharedString = Convert.ToInt32(cell.InnerText);
                        string? content = sheet.GetSharedString(indexSharedString);
                        Assert.That(value.ToString(), Is.EqualTo(content));
                    }
                    else if (BuildHelpers.DateTimeTypes.Contains(value.GetType()))
                    {
                        var date = ((DateTime)value).ToOADate();
                        double content = Convert.ToDouble(cell.CellValue!.Text, CultureInfo.InvariantCulture);
                        Assert.That(cell.DataType!.Value, Is.EqualTo(DocumentFormat.OpenXml.Spreadsheet.CellValues.Number));
                        Assert.That(date, Is.EqualTo(content));
                    }
                    else if (BuildHelpers.TimeSpanTypes.Contains(value.GetType()))
                    {
                        var totalDays = ((TimeSpan)value).TotalDays;
                        double content = Convert.ToDouble(cell.CellValue!.Text, CultureInfo.InvariantCulture);
                        Assert.That(cell.DataType!.Value, Is.EqualTo(DocumentFormat.OpenXml.Spreadsheet.CellValues.Number));
                        Assert.That(totalDays, Is.EqualTo(content));
                    }
                    else if (cell.DataType!.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.Number)
                    {
                        // Treat all as double
                        double content = Convert.ToDouble(cell.CellValue!.Text, CultureInfo.InvariantCulture);
                        double valueDouble = Convert.ToDouble(value);
                        Assert.That(valueDouble, Is.EqualTo(content));
                    }
                });
            }
        }

        internal static void AssertDimensions<T>(TabularSheet<T> table, SheetFixture sheet)
        {
            int rowCount = table.Count;
            int colCount = table.ColumnCount;

            string from = CellRefBuilder.BuildRef(0, 0);
            string to = CellRefBuilder.BuildRef(rowCount, colCount, false); //Non zero based index because is count
            string dimensions = CellRefBuilder.BuildRefRange(from, to);

            Assert.That(sheet.GetDimensionReference(), Is.EqualTo(dimensions));
        }
    }
}
