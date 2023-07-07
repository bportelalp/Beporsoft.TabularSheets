using Beporsoft.TabularSheets.Builders.SheetBuilders;
using Beporsoft.TabularSheets.CellStyling;
using Beporsoft.TabularSheets.Test.Helpers;
using System.Drawing;
using System.Globalization;
using System.Linq;
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
            Color bgColor = Color.Azure;
            double fontSize = 9.25;
            BorderStyle.BorderType borderType = BorderStyle.BorderType.Dashed;

            string path = GetPath($"Test{nameof(TryHeaderStyles)}.xlsx");
            TabularSheet<Product> table = null!;
            SheetWrapper sheet = null!;
            Assert.That(() =>
            {
                table = Generate();
                table.HeaderStyle.Fill.BackgroundColor = bgColor;
                table.HeaderStyle.Font.Size = fontSize;
                table.HeaderStyle.Border.SetBorderType(borderType);
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
                    Assert.That(style.Fill.BackgroundColor?.ToArgb(), Is.EqualTo(bgColor.ToArgb()));
                    Assert.That(style.Font.Size, Is.EqualTo(fontSize));
                    Assert.That(style.Border.Top, Is.EqualTo(borderType));
                    Assert.That(style.Border.Bottom, Is.EqualTo(borderType));
                    Assert.That(style.Border.Left, Is.EqualTo(borderType));
                    Assert.That(style.Border.Right, Is.EqualTo(borderType));
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
                    if (value.GetType() == typeof(DateTime) || value.GetType() == typeof(TimeSpan))
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
        public void TryBodyStyle()
        {
            Color bgColor = Color.Azure;
            string font = "Arial";
            double fontSize = 8;
            BorderStyle.BorderType borderType = BorderStyle.BorderType.Thin;

            string path = GetPath($"Test{nameof(TryBodyStyle)}.xlsx");
            TabularSheet<Product> table = null!;
            SheetWrapper sheet = null!;
            Assert.That(() =>
            {
                table = Generate();
                table.BodyStyle.Fill.BackgroundColor = bgColor;
                table.BodyStyle.Font.Font = font;
                table.BodyStyle.Font.Size = fontSize;
                table.BodyStyle.Border.SetBorderType(borderType);
                table.Create(path);
                sheet = new SheetWrapper(path);
            }, Throws.Nothing);
            AssertTabularSheetData(table, sheet);
            // Header default style
            foreach (var cell in sheet.GetHeaderCells())
            {
                if (cell.StyleIndex is not null)
                {
                    Style style = sheet.GetCellStyle(Convert.ToInt32(cell.StyleIndex.Value))!;
                    Assert.Multiple(() =>
                    {
                        Assert.That(style.Fill, Is.EqualTo(FillStyle.Default));
                        Assert.That(style.Font, Is.EqualTo(FontStyle.Default));
                        Assert.That(style.Border, Is.EqualTo(BorderStyle.Default));
                    });
                }
            }

            // Body has the specified style, Not check for numbering for datetime. Assumed ok due the previous test
            foreach (var cell in sheet.GetBodyCells())
            {
                Assert.That(cell.StyleIndex, Is.Not.Null);
                Style style = sheet.GetCellStyle(Convert.ToInt32(cell.StyleIndex.Value))!;
                Assert.Multiple(() =>
                {
                    Assert.That(style.Fill.BackgroundColor?.ToArgb(), Is.EqualTo(bgColor.ToArgb()));
                    Assert.That(style.Font.Size, Is.EqualTo(fontSize));
                    Assert.That(style.Font.Font, Is.EqualTo(font));
                    Assert.That(style.Border.Top, Is.EqualTo(borderType));
                    Assert.That(style.Border.Bottom, Is.EqualTo(borderType));
                    Assert.That(style.Border.Left, Is.EqualTo(borderType));
                    Assert.That(style.Border.Right, Is.EqualTo(borderType));
                });
            }

        }

        [Test]
        public void TryOverrideStyle()
        {
            Color bgColorHead = Color.DarkBlue;
            Color bgColorBody = Color.Azure;
            string fontBody = "Arial";
            Color fontColorHeader = Color.White;
            int fontSizeBody = 8;
            BorderStyle.BorderType borderTypeBody = BorderStyle.BorderType.Thin;
            BorderStyle.BorderType borderTypeHead = BorderStyle.BorderType.Medium;
            Color borderColorHead = Color.Yellow;

            string path = GetPath($"Test{nameof(TryOverrideStyle)}.xlsx");
            TabularSheet<Product> table = null!;
            SheetWrapper sheet = null!;
            Assert.That(() =>
            {
                table = Generate();
                table.BodyStyle.Fill.BackgroundColor = bgColorBody;
                table.BodyStyle.Font.Font = fontBody;
                table.BodyStyle.Font.Size = fontSizeBody;
                table.BodyStyle.Border.SetBorderType(borderTypeBody);

                table.HeaderStyle.Fill.BackgroundColor = bgColorHead;
                table.HeaderStyle.Font.Color = fontColorHeader;
                table.HeaderStyle.Border.SetBorderType(borderTypeHead);
                table.HeaderStyle.Border.Color = borderColorHead;

                table.Options.InheritHeaderStyleFromBody = true;
                table.Create(path);
                sheet = new SheetWrapper(path);
            }, Throws.Nothing);
            AssertTabularSheetData(table, sheet);
            // Header merged style
            foreach (var cell in sheet.GetHeaderCells())
            {
                Assert.That(cell.StyleIndex, Is.Not.Null);
                Style style = sheet.GetCellStyle(Convert.ToInt32(cell.StyleIndex.Value))!;
                Assert.Multiple(() =>
                {
                    Assert.That(style.Fill.BackgroundColor?.ToArgb(), Is.EqualTo(bgColorHead.ToArgb()));
                    Assert.That(style.Font.Font, Is.EqualTo(fontBody));
                    Assert.That(style.Font.Color?.ToArgb(), Is.EqualTo(fontColorHeader.ToArgb()));
                    Assert.That(style.Font.Size, Is.EqualTo(fontSizeBody));
                    Assert.That(style.Border.Color?.ToArgb(), Is.EqualTo(borderColorHead.ToArgb()));
                    Assert.That(style.Border.Top, Is.EqualTo(borderTypeHead));
                    Assert.That(style.Border.Bottom, Is.EqualTo(borderTypeHead));
                    Assert.That(style.Border.Left, Is.EqualTo(borderTypeHead));
                    Assert.That(style.Border.Right, Is.EqualTo(borderTypeHead));
                });

            }
            // check defaults body and not inherited from header
            foreach (var cell in sheet.GetBodyCells())
            {
                Assert.That(cell.StyleIndex, Is.Not.Null);
                Style style = sheet.GetCellStyle(Convert.ToInt32(cell.StyleIndex.Value))!;
                Assert.Multiple(() =>
                {
                    Assert.That(style.Fill.BackgroundColor?.ToArgb(), Is.EqualTo(bgColorBody.ToArgb()));
                    Assert.That(style.Font.Font, Is.EqualTo(fontBody));
                    Assert.That(style.Font.Size, Is.EqualTo(fontSizeBody));
                    Assert.That(style.Font.Color?.ToArgb(), Is.Not.EqualTo(fontColorHeader.ToArgb()));
                    Assert.That(style.Border.Color?.ToArgb(), Is.Not.EqualTo(borderColorHead.ToArgb()));
                    Assert.That(style.Border.Top, Is.EqualTo(borderTypeBody));
                    Assert.That(style.Border.Bottom, Is.EqualTo(borderTypeBody));
                    Assert.That(style.Border.Left, Is.EqualTo(borderTypeBody));
                    Assert.That(style.Border.Right, Is.EqualTo(borderTypeBody));
                });

            }
        }

        [Test]
        public void TryColumnStyle()
        {
            string path = GetPath($"Test{nameof(TryColumnStyle)}.xlsx");
            TabularSheet<Product> table = null!;
            SheetWrapper sheet = null!;
            int indexColExtra1 = 0;
            int indexColExtra2 = 0;
            int indexColExtra3 = 0;
            int indexColExtra4 = 0;
            Assert.That(() =>
            {
                Style style = new Style();
                style.Fill.BackgroundColor = Color.Red;
                table = Generate();
                var colExtra1 = table.AddColumn(t => t.Name).SetTitle("Name with new style").SetStyle(style);
                var colExtra2 = table.AddColumn(t => t.Cost).SetTitle("Cost 2 decimals").SetStyle(s =>
                {
                    s.NumberingPattern = "0.00";
                    s.Fill.BackgroundColor = Color.AliceBlue;
                });
                indexColExtra1 = colExtra1.ColumnIndex;
                indexColExtra2 = colExtra2.ColumnIndex;

                var colExtra3 = table.AddColumn(t => t.LastPriceUpdate)
                                        .SetTitle("Data unmodify Numbering")
                                        .SetStyle(s => s.Font.Color = Color.Blue);
                var colExtra4 = table.AddColumn(t => t.LastPriceUpdate)
                                        .SetTitle("Data modify Numbering")
                                        .SetStyle(s => s.NumberingPattern = "d-mmm-yy");
                indexColExtra3 = colExtra3.ColumnIndex;
                indexColExtra4 = colExtra4.ColumnIndex;

                string path = GetPath($"Test{nameof(TryColumnStyle)}.xlsx");
                table.Create(path);

                sheet = new SheetWrapper(path);
            }, Throws.Nothing);

            AssertTabularSheetData(table, sheet);
            foreach (var cell in sheet.GetBodyCellsByColumn(indexColExtra1))
            {
                Assert.That(cell.StyleIndex, Is.Not.Null);
                Style style = sheet.GetCellStyle(Convert.ToInt32(cell.StyleIndex.Value))!;
                Assert.That(style.Fill.BackgroundColor?.ToArgb(), Is.EqualTo(Color.Red.ToArgb()));
            }
            foreach (var cell in sheet.GetBodyCellsByColumn(indexColExtra2))
            {
                Assert.That(cell.StyleIndex, Is.Not.Null);
                Style style = sheet.GetCellStyle(Convert.ToInt32(cell.StyleIndex.Value))!;
                Assert.Multiple(() =>
                {
                    Assert.That(style.Fill.BackgroundColor?.ToArgb(), Is.EqualTo(Color.AliceBlue.ToArgb()));
                    Assert.That(style.NumberingPattern, Is.EqualTo("0.00"));
                });
            }

            foreach (var cell in sheet.GetBodyCellsByColumn(indexColExtra3))
            {
                Assert.That(cell.StyleIndex, Is.Not.Null);
                Style style = sheet.GetCellStyle(Convert.ToInt32(cell.StyleIndex.Value))!;
                Assert.Multiple(() =>
                {
                    Assert.That(style.Font.Color?.ToArgb(), Is.EqualTo(Color.Blue.ToArgb()));
                    Assert.That(style.NumberingPattern, Is.EqualTo(table.Options.DateTimeFormat));
                });
            }

            foreach (var cell in sheet.GetBodyCellsByColumn(indexColExtra4))
            {
                Assert.That(cell.StyleIndex, Is.Not.Null);
                Style style = sheet.GetCellStyle(Convert.ToInt32(cell.StyleIndex.Value))!;
                Assert.That(style.NumberingPattern, Is.EqualTo("d-mmm-yy"));
            }
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
                    else if (CellBuilder.DateTimeTypes.Contains(value.GetType()))
                    {
                        var date = ((DateTime)value).ToOADate();
                        double content = Convert.ToDouble(cell.CellValue!.Text, CultureInfo.InvariantCulture);
                        Assert.That(cell.DataType!.Value, Is.EqualTo(DocumentFormat.OpenXml.Spreadsheet.CellValues.Number));
                        Assert.That(date, Is.EqualTo(content));
                    }
                    else if (CellBuilder.TimeSpanTypes.Contains(value.GetType()))
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
            table.AddColumn(t => t.DeliveryTime).SetTitle(nameof(Product.DeliveryTime));
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