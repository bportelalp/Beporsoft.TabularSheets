using Beporsoft.TabularSheets.Builders.SheetBuilders;
using Beporsoft.TabularSheets.CellStyling;
using Beporsoft.TabularSheets.Test.Helpers;
using System.Drawing;
using System.Globalization;
using System.Reflection;

namespace Beporsoft.TabularSheets.Test.TestsTabularData
{
    public class TestTabularSheet
    {

        [Test]
        public void CheckFileName()
        {
            TabularSheet<Product> table = Generate();
            string pathOk = GetPath("ExcelCheckFileName.xlsx");
            string pathOkAlternative = GetPath("ExcelCheckFileName.xls");
            string pathWrongExtension = GetPath("ExcelCheckFileName.csv");
            Assert.Multiple(() =>
            {
                Assert.That(() => table.Create(pathOk), Throws.Nothing);
                Assert.That(() => table.Create(pathOkAlternative), Throws.Nothing);
                Assert.That(() => table.Create(pathWrongExtension), Throws.Exception);
            });
        }

        [Test]
        public void TestToDebug()
        {
            Assert.That(() =>
            {
                TabularSheet<Product> table = Generate();
                table.HeaderStyle.Fill.BackgroundColor = Color.Purple;
                table.BodyStyle.Font.Color = Color.Red;
                table.HeaderStyle.Font.Color = Color.White;
                table.BodyStyle.Fill.BackgroundColor = Color.AliceBlue;
                table.BodyStyle.Border.SetBorderType(BorderStyle.BorderType.Thin);
                string path = GetPath("DebugTest.xlsx");
                table.Create(path);
            }, Throws.Nothing);
        }

        [Test]
        public void TryDataIntegrity()
        {
            string path = GetPath($"Test{nameof(TryDataIntegrity)}.xlsx");
            TabularSheet<Product> table = null!;
            SheetWrapper sheet = null!;
            Assert.That(() =>
            {
                table = Generate();
                table.Create(path);
                sheet = new SheetWrapper(path);
            }, Throws.Nothing);

            AssertTabularSheetData(table, sheet);
        }

        [Test]
        public void TryHeaderStyles()
        {
            string path = GetPath($"Test{nameof(TryHeaderStyles)}.xlsx");
            TabularSheet<Product> table = null!;
            SheetWrapper sheet = null!;
            Assert.That(() =>
            {
                table = Generate();
                table.HeaderStyle.Fill.BackgroundColor = Color.Azure;
                table.HeaderStyle.Font.Size = 12;
                table.HeaderStyle.Border.SetBorderType(BorderStyle.BorderType.Medium);
                table.Create(path);
                sheet = new SheetWrapper(path);
            }, Throws.Nothing);
            AssertTabularSheetData(table, sheet);

            foreach (var cell in sheet.GetHeaderCells())
            {
                Assert.That(cell.StyleIndex, Is.Not.Null);
                Style? style = sheet.GetCellStyle(Convert.ToInt32(cell.StyleIndex.Value));
                Assert.That(style, Is.Not.Null);
                Assert.Multiple(() =>
                {
                    Assert.That(style.Fill.BackgroundColor?.ToArgb(), Is.EqualTo(Color.Azure.ToArgb()));
                    Assert.That(style.Font.Size, Is.EqualTo(12));
                });
            }
            foreach (var cell in sheet.GetBodyCells())
            {
                if (cell.StyleIndex is not null)
                {
                    Style style = sheet.GetCellStyle(Convert.ToInt32(cell.StyleIndex.Value))!;
                    // Styles must be the default except datetime which has numbering pattern
                    Assert.Multiple(() =>
                    {
                        Assert.That(style.Fill, Is.EqualTo(FillStyle.Default));
                        Assert.That(style.Font, Is.EqualTo(FontStyle.Default));
                        Assert.That(style.Border, Is.EqualTo(BorderStyle.Default));
                    });

                    (int row, int col) = CellRefBuilder.GetIndexes(cell.CellReference!.Value!);
                    object value = table.Columns.Single(c => c.ColumnIndex == col).Apply(table.Items[row - 1]);
                    if (value.GetType() == typeof(DateTime))
                    {
                        Assert.That(style.NumberingPattern, Is.Not.Null);
                        Assert.That(style.NumberingPattern, Is.Not.Empty);
                    }
                    else
                    {
                        Assert.That(style.NumberingPattern, Is.Null);
                    }
                }
            }
        }

        [Test]
        public void TryDefaultStyles()
        {
            Assert.That(() =>
            {
                TabularSheet<Product> table = Generate();
                table.BodyStyle.Fill.BackgroundColor = Color.Azure;
                table.BodyStyle.Font.Font = "Arial";
                table.BodyStyle.Font.Size = 8;
                table.BodyStyle.Border.SetBorderType(BorderStyle.BorderType.Thin);
                string path = GetPath($"Test{nameof(TryDefaultStyles)}.xlsx");
                table.Create(path);
            }, Throws.Nothing);
        }

        [Test]
        public void TryOverrideStyle()
        {
            Assert.That(() =>
            {
                TabularSheet<Product> table = Generate();
                table.BodyStyle.Fill.BackgroundColor = Color.Azure;
                table.BodyStyle.Font.Font = "Arial";
                table.BodyStyle.Font.Size = 8;
                table.BodyStyle.Border.SetBorderType(BorderStyle.BorderType.Thin);

                table.HeaderStyle.Fill.BackgroundColor = Color.DarkBlue;
                table.HeaderStyle.Font.Color = Color.White;
                table.HeaderStyle.Border.SetBorderType(BorderStyle.BorderType.Medium);
                table.HeaderStyle.Border.Color = Color.Yellow;
                string path = GetPath($"Test{nameof(TryOverrideStyle)}.xlsx");
                table.Create(path);
            }, Throws.Nothing);
        }

        [Test]
        public void TryColumnStyle()
        {
            Assert.That(() =>
            {
                Style style = new Style();
                style.Fill.BackgroundColor = Color.Red;
                TabularSheet<Product> table = Generate();
                table.AddColumn(t => t.Name).SetTitle("Name with new style").SetStyle(style);
                table.AddColumn(t => t.Cost).SetTitle("Cost 2 decimals").SetStyle(s =>
                {
                    s.NumberingPattern = "0.00";
                    s.Fill.BackgroundColor = Color.AliceBlue;
                });
                table.AddColumn(t => t.LastPriceUpdate).SetTitle("Data unmodify Numbering").SetStyle(s => s.Font.Color = Color.Blue);
                table.AddColumn(t => t.LastPriceUpdate).SetTitle("Data modify Numbering").SetStyle(s => s.NumberingPattern = "d-mmm-yy");
                string path = GetPath($"Test{nameof(TryColumnStyle)}.xlsx");
                table.Create(path);
            }, Throws.Nothing);
        }

        #region TestHelpers
        /// <summary>
        /// Verify the data on every column is OK as expected
        /// </summary>
        /// <param name="table"></param>
        /// <param name="column"></param>
        /// <param name="sheet"></param>
        private static void AssertTabularSheetData(TabularSheet<Product> table, SheetWrapper sheet)
        {
            foreach (var col in table.Columns)
            {
                AssertColumnHeaderData(col, sheet);
                AssertColumnBodyData(table, col, sheet);
            }

        }

        private static void AssertColumnHeaderData(TabularDataColumn<Product> column, SheetWrapper sheet)
        {
            DocumentFormat.OpenXml.Spreadsheet.Cell headerCell = sheet.GetHeaderCellByColumn(column.ColumnIndex);
            Assert.Multiple(() =>
            {
                Assert.That(headerCell.InnerText, Is.Not.Null);
                Assert.That(headerCell.DataType!.Value, Is.EqualTo(DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString));
                var indexSharedString = Convert.ToInt32(headerCell.InnerText);
                string? headerTitle = sheet.GetSharedString(indexSharedString);
                Assert.That(headerTitle, Is.EqualTo(column.Title));
            });
        }

        private static void AssertColumnBodyData(TabularSheet<Product> table, TabularDataColumn<Product> column, SheetWrapper sheet)
        {
            List<DocumentFormat.OpenXml.Spreadsheet.Cell> bodyCells = sheet.GetBodyCellsByColumn(column.ColumnIndex);
            foreach (var cell in bodyCells)
            {
                Assert.Multiple(() =>
                {
                    int row = CellRefBuilder.GetRowIndex(cell.CellReference!);
                    Product item = table[row - 1];
                    object value = column.Apply(item);
                    if (cell.DataType!.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString)
                    {
                        var indexSharedString = Convert.ToInt32(cell.InnerText);
                        string? content = sheet.GetSharedString(indexSharedString);
                        Assert.That(value.ToString(), Is.EqualTo(content));
                    }
                    else if (value.GetType() == typeof(DateTime))
                    {
                        var date = ((DateTime)value).ToOADate();
                        double content = Convert.ToDouble(cell.CellValue!.Text);
                        Assert.That(cell.DataType!.Value, Is.EqualTo(DocumentFormat.OpenXml.Spreadsheet.CellValues.Number));
                        Assert.That(date, Is.EqualTo(content));
                    }
                    else if (cell.DataType!.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.Number)
                    {
                        // Treat all as double
                        double content = Convert.ToDouble(cell.CellValue!.Text, System.Globalization.CultureInfo.InvariantCulture);
                        double valueDouble = Convert.ToDouble(column.Apply(item));
                        Assert.That(valueDouble, Is.EqualTo(content));
                    }
                });
            }
        }
        #endregion

        #region Data Helpers
        private static TabularSheet<Product> Generate()
        {
            TabularSheet<Product> table = new();
            table.AddRange(Product.GenerateProducts(50));

            table.AddColumn(t => t.Id).SetTitle(nameof(Product.Id));
            table.AddColumn(t => t.Name).SetTitle(nameof(Product.Name));
            table.AddColumn(t => t.Vendor).SetTitle(nameof(Product.Vendor));
            table.AddColumn(t => t.CountryOrigin).SetTitle(nameof(Product.CountryOrigin));
            table.AddColumn(nameof(Product.Cost), t => t.Cost);
            table.AddColumn(t => t.LastPriceUpdate).SetTitle(nameof(Product.LastPriceUpdate));
            table.AddColumn(t => t.LastUpdate).SetTitle(nameof(Product.LastUpdate)); ;
            return table;
        }


        private string GetPath(string fileName)
        {
            DirectoryInfo? projectDir = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory)?.Parent?.Parent?.Parent;
            return $"{projectDir!.FullName}\\Results\\{fileName}";
        }
        #endregion

    }
}