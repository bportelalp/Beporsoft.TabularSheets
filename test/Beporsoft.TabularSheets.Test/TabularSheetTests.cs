using Beporsoft.TabularSheets.Builders;
using Beporsoft.TabularSheets.Builders.SheetBuilders;
using Beporsoft.TabularSheets.CellStyling;
using Beporsoft.TabularSheets.Options;
using Beporsoft.TabularSheets.Options.ColumnWidth;
using Beporsoft.TabularSheets.Test.Helpers;
using Beporsoft.TabularSheets.Test.TestModels;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;

namespace Beporsoft.TabularSheets.Test
{
    [Category("SheetStructure")]
    public class TabularSheetTests
    {
        private readonly bool _clearFolderOnEnd = false;
        private readonly int _amountRows = 1000;
        private readonly TestFilesHandler _filesHandler = new(nameof(TabularSheetTests));
        private readonly Stopwatch _stopwatch = new Stopwatch();

        [SetUp]
        public void Setup()
        {
            _stopwatch.Reset();
        }

        [Test, Category("DEBUG")]
        public void DebugFile()
        {
            Assert.That(() =>
            {
                TabularSheet<Product> table = Product.GenerateProductSheet(_amountRows);
                table.HeaderStyle.Fill.BackgroundColor = Color.Purple;
                table.BodyStyle.Font.Color = Color.Red;
                table.HeaderStyle.Font.Color = Color.White;
                table.BodyStyle.Fill.BackgroundColor = Color.AliceBlue;
                table.BodyStyle.Border.SetBorderType(BorderStyle.BorderType.Thin);
                table.Options.ColumnOptions.Width = new AutoColumnWidth();
                string path = _filesHandler.BuildPath($"Test{nameof(DebugFile)}.xlsx");
                table.Columns.Single(c => c.Index == 0).Options.Width = new FixedColumnWidth(10);
                table.Create(path);
            }, Throws.Nothing);
        }

        [Test, Category("Cells"), Category("Worksheet")]
        public void Create_Data_AsExpected()
        {
            string path = _filesHandler.BuildPath($"Test{nameof(Create_Data_AsExpected)}.xlsx");
            TabularSheet<Product> table = null!;
            SheetFixture sheet = null!;
            // Generate as file
            Assert.That(() =>
            {
                table = Product.GenerateProductSheet();
                table.Create(path);
                sheet = new SheetFixture(path);
            }, Throws.Nothing);
            TabularSheetAsserter.AssertTabularSheet(table, sheet);

            // Generate as memory stream
            Assert.That(() =>
            {
                table = Product.GenerateProductSheet();
                _stopwatch.Start();
                using MemoryStream ms = table.Create();
                Console.WriteLine($"Created on: {_stopwatch.Elapsed.TotalMilliseconds:F3}ms");
                sheet = new SheetFixture(ms);
            }, Throws.Nothing);
            TabularSheetAsserter.AssertTabularSheet(table, sheet);
        }

        [Test, Category("Stylesheet"), Category("Worksheet")]
        public void Create_HeaderStyles_AsExpected()
        {
            Color bgColor = Color.Azure;
            double fontSize = 11.25;
            BorderStyle.BorderType borderType = BorderStyle.BorderType.Dashed;
            bool bold = true;
            bool italic = true;

            string path = _filesHandler.BuildPath($"Test{nameof(Create_HeaderStyles_AsExpected)}.xlsx");
            TabularSheet<Product> table = null!;
            SheetFixture sheet = null!;
            Assert.That(() =>
            {
                table = Product.GenerateProductSheet(_amountRows); 
                table.HeaderStyle.Fill.BackgroundColor = bgColor;
                table.HeaderStyle.Font.Size = fontSize;
                table.HeaderStyle.Font.Bold = bold;
                table.HeaderStyle.Font.Italic = italic;
                table.HeaderStyle.Border.SetBorderType(borderType);
                _stopwatch.Start();
                table.Create(path);
                Console.WriteLine($"Created on: {_stopwatch.Elapsed.TotalMilliseconds:F3}ms");
                sheet = new SheetFixture(path);
            }, Throws.Nothing);

            foreach (var cell in sheet.GetHeaderCells())
            {
                Assert.That(cell.StyleIndex, Is.Not.Null);
                Style? style = sheet.GetCellStyle(Convert.ToInt32(cell.StyleIndex.Value));
                Assert.That(style, Is.Not.Null);
                Assert.Multiple(() =>
                {
                    Assert.That(style.Fill.BackgroundColor?.ToArgb(), Is.EqualTo(bgColor.ToArgb()));
                    Assert.That(style.Font.Size, Is.EqualTo(fontSize));
                    Assert.That(style.Font.Bold, Is.EqualTo(bold));
                    Assert.That(style.Font.Italic, Is.EqualTo(italic));
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
                    object value = table.Columns.Single(c => c.Index == col).Apply(table.Items[row - 1]);
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

        [Test, Category("Stylesheet"), Category("Worksheet")]
        public void Create_BodyStyles_AsExpected()
        {
            Color bgColor = Color.Azure;
            string font = "Arial";
            double fontSize = 8;
            BorderStyle.BorderType borderType = BorderStyle.BorderType.Thin;

            string path = _filesHandler.BuildPath($"Test{nameof(Create_BodyStyles_AsExpected)}.xlsx");
            TabularSheet<Product> table = null!;
            SheetFixture sheet = null!;
            Assert.That(() =>
            {
                table = Product.GenerateProductSheet(_amountRows);
                table.BodyStyle.Fill.BackgroundColor = bgColor;
                table.BodyStyle.Font.FontName = font;
                table.BodyStyle.Font.Size = fontSize;
                table.BodyStyle.Border.SetBorderType(borderType);
                _stopwatch.Start();
                table.Create(path);
                Console.WriteLine($"Created on: {_stopwatch.Elapsed.TotalMilliseconds:F3}ms");
                sheet = new SheetFixture(path);
            }, Throws.Nothing);

            // Header default style
            foreach (var cell in sheet.GetHeaderCells())
            {
                if (cell.StyleIndex is not null)
                {
                    Style style = sheet.GetCellStyle(Convert.ToInt32(cell.StyleIndex.Value))!;
                    Assert.Multiple(() =>
                    {
                        Assert.That(style.Fill, Is.EqualTo(FillStyle.Default));
                        Assert.That(style.Font, Is.EqualTo(CellStyling.FontStyle.Default));
                        Assert.That(style.Border, Is.EqualTo(BorderStyle.Default));
                    });
                }
            }

            //Body has the specified style, Not check for numbering for datetime.Assumed ok due the previous test
            foreach (var cell in sheet.GetBodyCells())
            {
                Assert.That(cell.StyleIndex, Is.Not.Null);
                Style style = sheet.GetCellStyle(Convert.ToInt32(cell.StyleIndex.Value))!;
                Assert.Multiple(() =>
                {
                    Assert.That(style.Fill.BackgroundColor?.ToArgb(), Is.EqualTo(bgColor.ToArgb()));
                    Assert.That(style.Font.Size, Is.EqualTo(fontSize));
                    Assert.That(style.Font.FontName, Is.EqualTo(font));
                    Assert.That(style.Border.Top, Is.EqualTo(borderType));
                    Assert.That(style.Border.Bottom, Is.EqualTo(borderType));
                    Assert.That(style.Border.Left, Is.EqualTo(borderType));
                    Assert.That(style.Border.Right, Is.EqualTo(borderType));
                });
            }
        }

        [Test, Category("Stylesheet"), Category("StyleCombination"), Category("Worksheet")]
        public void Create_OverrideStyles_DoneAsPreferences()
        {
            Color bgColorHead = Color.DarkBlue;
            Color bgColorBody = Color.Azure;
            string fontBody = "Arial";
            Color fontColorHeader = Color.White;
            int fontSizeBody = 8;
            BorderStyle.BorderType borderTypeBody = BorderStyle.BorderType.Thin;
            BorderStyle.BorderType borderTypeHead = BorderStyle.BorderType.Medium;
            Color borderColorHead = Color.Coral;
            bool inheritHeaderFromBody = true;

            string path = _filesHandler.BuildPath($"Test{nameof(Create_OverrideStyles_DoneAsPreferences)}.xlsx");
            TabularSheet<Product> table = null!;
            SheetFixture sheet = null!;
            Assert.That(() =>
            {
                table = Product.GenerateProductSheet(_amountRows);
                table.BodyStyle.Fill.BackgroundColor = bgColorBody;
                table.BodyStyle.Font.FontName = fontBody;
                table.BodyStyle.Font.Size = fontSizeBody;
                table.BodyStyle.Border.SetBorderType(borderTypeBody);

                table.HeaderStyle.Fill.BackgroundColor = bgColorHead;
                table.HeaderStyle.Font.Color = fontColorHeader;
                table.HeaderStyle.Border.SetBorderType(borderTypeHead);
                table.HeaderStyle.Border.Color = borderColorHead;

                table.Options.InheritHeaderStyleFromBody = inheritHeaderFromBody;
                _stopwatch.Start();
                table.Create(path);
                Console.WriteLine($"Created on: {_stopwatch.Elapsed.TotalMilliseconds:F3}ms");
                sheet = new SheetFixture(path);
            }, Throws.Nothing);

            // Header merged style
            foreach (var cell in sheet.GetHeaderCells())
            {
                Assert.That(cell.StyleIndex, Is.Not.Null);
                Style style = sheet.GetCellStyle(Convert.ToInt32(cell.StyleIndex.Value))!;
                Assert.Multiple(() =>
                {
                    Assert.That(style.Fill.BackgroundColor?.ToArgb(), Is.EqualTo(bgColorHead.ToArgb()));
                    Assert.That(style.Font.FontName, Is.EqualTo(fontBody));
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
                    Assert.That(style.Font.FontName, Is.EqualTo(fontBody));
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

        [Test, Category("Stylesheet"), Category("StyleCombination"), Category("Worksheet")]
        public void Create_ColumnStyles_AsExpected()
        {
            string path = _filesHandler.BuildPath($"Test{nameof(Create_ColumnStyles_AsExpected)}.xlsx");
            TabularSheet<Product> table = null!;
            SheetFixture sheet = null!;
            int indexColExtra1 = 0;
            int indexColExtra2 = 0;
            int indexColExtra3 = 0;
            int indexColExtra4 = 0;
            Assert.That(() =>
            {
                Style style = new();
                style.Fill.BackgroundColor = Color.Red;
                table = Product.GenerateProductSheet(_amountRows);
                var colExtra1 = table.AddColumn(t => t.Name).SetTitle("Name with new style").SetStyle(style);
                var colExtra2 = table.AddColumn(t => t.Cost).SetTitle("Cost 2 decimals").SetStyle(s =>
                {
                    s.NumberingPattern = "0.00";
                    s.Fill.BackgroundColor = Color.AliceBlue;
                    s.Font.Underline = FontStyle.UnderlineType.Single;
                });
                indexColExtra1 = colExtra1.Index;
                indexColExtra2 = colExtra2.Index;

                var colExtra3 = table.AddColumn(t => t.LastPriceUpdate)
                                        .SetTitle("Data unmodify Numbering")
                                        .SetStyle(s => s.Font.Color = Color.Blue);
                var colExtra4 = table.AddColumn(t => t.LastPriceUpdate)
                                        .SetTitle("Data modify Numbering")
                                        .SetStyle(s => s.NumberingPattern = "d-mmm-yy");
                indexColExtra3 = colExtra3.Index;
                indexColExtra4 = colExtra4.Index;

                string path = _filesHandler.BuildPath($"Test{nameof(Create_ColumnStyles_AsExpected)}.xlsx");
                _stopwatch.Start();
                table.Create(path);
                Console.WriteLine($"Created on: {_stopwatch.Elapsed.TotalMilliseconds:F3}ms");

                sheet = new SheetFixture(path);
            }, Throws.Nothing);

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
                    Assert.That(style.Font.Underline, Is.EqualTo(FontStyle.UnderlineType.Single));
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

        [Test, Category("Stylesheet"), Category("Worksheet")]
        public void Create_AligmentStyles_AsExpected()
        {
            AlignmentStyle.HorizontalAlignment horizontalCol0 = AlignmentStyle.HorizontalAlignment.Center;
            AlignmentStyle.VerticalAlignment verticalCol0 = AlignmentStyle.VerticalAlignment.Center;
            bool textWrapCol0 = true;

            string path = _filesHandler.BuildPath($"Test{nameof(Create_AligmentStyles_AsExpected)}.xlsx");
            TabularSheet<Product> table = null!;
            SheetFixture sheet = null!;
            Assert.That(() =>
            {
                table = Product.GenerateProductSheet(_amountRows);
                var col0 = table.Columns.Single(c => c.Index == 0);
                col0.SetStyle(s =>
                {
                    s.Alignment.Horizontal = horizontalCol0;
                    s.Alignment.Vertical = verticalCol0;
                    s.Alignment.TextWrap = textWrapCol0;
                });
                _stopwatch.Start();
                table.Create(path);
                Console.WriteLine($"Created on: {_stopwatch.Elapsed.TotalMilliseconds:F3}ms");
                sheet = new SheetFixture(path);
            }, Throws.Nothing);

            foreach (var cell in sheet.GetBodyCellsByColumn(0))
            {
                Assert.That(cell.StyleIndex, Is.Not.Null);
                Style style = sheet.GetCellStyle(Convert.ToInt32(cell.StyleIndex.Value))!;
                Assert.Multiple(() =>
                {
                    Assert.That(style, Is.Not.Null);
                    Assert.That(style.Alignment.Vertical, Is.EqualTo(verticalCol0));
                    Assert.That(style.Alignment.Horizontal, Is.EqualTo(horizontalCol0));
                    Assert.That(style.Alignment.TextWrap, Is.EqualTo(textWrapCol0));
                });

            }

        }

        [Test, Category("Stylesheet"), Category("Worksheet"), Category("Columns")]
        public void Create_ColumnWidth_AsExpected()
        {
            IColumnWidth? tableWidth = new AutoColumnWidth(1.1);
            string fontName = "Arial";
            double fontSize = 15;


            string path = _filesHandler.BuildPath($"Test{nameof(Create_ColumnWidth_AsExpected)}.xlsx");
            TabularSheet<Product> table = null!;
            SheetFixture sheet = null!;
            Assert.That(() =>
            {
                table = Product.GenerateProductSheet(_amountRows);
                table.BodyStyle.Font.FontName = fontName;
                table.BodyStyle.Font.Size = fontSize;
                table.Options.InheritHeaderStyleFromBody = true;
                table.Options.ColumnOptions.Width = tableWidth;
                table.HeaderStyle.Fill.BackgroundColor = Color.Azure;
                table.Columns.Single(c => c.Index == 4).Style.NumberingPattern = "0.0000000000000";
                _stopwatch.Start();
                table.Create(path);
                Console.WriteLine($"Created on: {_stopwatch.Elapsed.TotalMilliseconds:F3}ms");
                sheet = new SheetFixture(path);
            }, Throws.Nothing);
        }


        [OneTimeTearDown]
        public void OneTimeTearDown()
        {
            if (_clearFolderOnEnd)
                _filesHandler.ClearFiles();
            
        }

    }
}