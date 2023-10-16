using Beporsoft.TabularSheets.Builders.Interfaces;
using Beporsoft.TabularSheets.Builders.SheetBuilders;
using Beporsoft.TabularSheets.Builders.SheetBuilders.Adapters;
using Beporsoft.TabularSheets.Builders.StyleBuilders;
using Beporsoft.TabularSheets.Builders.StyleBuilders.Adapters;
using Beporsoft.TabularSheets.CellStyling;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using System.Diagnostics;
using System.Drawing;

namespace Beporsoft.TabularSheets.Test
{
    [Category("Stylesheet")]
    internal class TestCellStyling
    {
        private const int _amountIndexedSetupItems = 10000;
        private readonly Stopwatch _stopwatch = new();

        [SetUp]
        public void Setup()
        {
            _stopwatch.Reset();
        }

        [Test, Category("StyleEquality")]
        public void BorderStyle_IEquatableCompliant()
        {
            var border = new BorderStyle();
            Assert.That(border, Is.EqualTo(BorderStyle.Default));

            var borderModified = border;
            border.SetBorderType(BorderStyle.BorderType.Thin);
            border.Color = System.Drawing.Color.Aquamarine;
            Assert.Multiple(() =>
            {
                Assert.That(borderModified, Is.EqualTo(border));
                Assert.That(borderModified, Is.Not.EqualTo(BorderStyle.Default));
                Assert.That(border, Is.Not.EqualTo(BorderStyle.Default));
            });

            border.SetBorderType(null);
            border.Color = null;
            Assert.Multiple(() =>
            {
                Assert.That(borderModified, Is.EqualTo(border));
                Assert.That(borderModified, Is.EqualTo(BorderStyle.Default));
                Assert.That(border, Is.EqualTo(BorderStyle.Default));
            });
        }

        [Test, Category("StyleEquality")]
        public void FontStyle_IEquatableCompliant()
        {
            var font = new FontStyle();
            Assert.That(font, Is.EqualTo(FontStyle.Default));

            font.Size = 12;
            font.FontName = "efa";
            font.Color = System.Drawing.Color.Aquamarine;
            Assert.That(font, Is.Not.EqualTo(FontStyle.Default));

            font.Size = null;
            font.FontName = null;
            font.Color = null;
            Assert.That(font, Is.EqualTo(FontStyle.Default));
        }

        [Test, Category("StyleEquality")]
        public void FillStyle_IEquatableCompliant()
        {
            var fill = new FillStyle();
            Assert.That(fill, Is.EqualTo(FillStyle.Default));

            fill.BackgroundColor = System.Drawing.Color.Aquamarine;
            Assert.That(fill, Is.Not.EqualTo(FillStyle.Default));

            fill.BackgroundColor = null;
            Assert.That(fill, Is.EqualTo(FillStyle.Default));
        }

        [Test, Category("StyleEquality")]
        public void Style_IEquatableCompliant()
        {
            var style = new Style();
            Assert.Multiple(() =>
            {
                Assert.That(style, Is.EqualTo(Style.Default));
                Assert.That(style.Font, Is.EqualTo(FontStyle.Default));
                Assert.That(style.Fill, Is.EqualTo(FillStyle.Default));
                Assert.That(style.Border, Is.EqualTo(BorderStyle.Default));
            });

            // Change one field of one property
            style.Font.Size = 0;
            Assert.Multiple(() =>
            {
                Assert.That(style, Is.Not.EqualTo(Style.Default));
                Assert.That(style.Font, Is.Not.EqualTo(FontStyle.Default));
                Assert.That(style.Fill, Is.EqualTo(FillStyle.Default));
                Assert.That(style.Border, Is.EqualTo(BorderStyle.Default));
            });
        }

        [Test, Category("StyleCombination")]
        [TestCaseSource(nameof(Data_CombineStylesCases))]
        public void StyleCombiner_Combine_IEquatableCompliant(Style style1, Style style2, Style styleExpected)
        {
            Style styleResult = StyleCombiner.Combine(style1, style2);
            Assert.That(styleResult, Is.EqualTo(styleExpected));
        }


        [Test, Category("StyleEquality"), Category("ISetupCollection")]
        [TestCaseSource(nameof(Data_SharedStringSetupCollection))]
        [TestCaseSource(nameof(Data_IndexedSetupCollection_Strings))]
        [TestCaseSource(nameof(Data_IndexedSetupCollection_BorderSetup))]
        [TestCaseSource(nameof(Data_IndexedSetupCollection_FillSetup))]
        [TestCaseSource(nameof(Data_IndexedSetupCollection_FontSetup))]
        [TestCaseSource(nameof(Data_IndexedSetupCollection_FormatSetup))]
        public void SetupCollection_Register_MatchEquivalent<TSetup>(List<TSetup> randomSetups, ISetupCollection<TSetup> collection)
            where TSetup : Setup, IEquatable<TSetup>, IIndexedSetup
        {
            List<TSetup> externalPlainCollection = new List<TSetup>();
            foreach (var setup in randomSetups)
            {
                if (!externalPlainCollection.Contains(setup))
                    externalPlainCollection.Add(setup);
                int indexExpected = externalPlainCollection.IndexOf(setup);

                _stopwatch.Start();
                int resultIndex = collection.Register(setup);
                _stopwatch.Stop();

                Assert.Multiple(() =>
                {
                    // Assert indexing on the right position
                    Assert.That(resultIndex, Is.EqualTo(indexExpected));
                    Assert.That(setup.Index, Is.EqualTo(indexExpected));
                    Assert.That(setup.Build(), Is.Not.Null);
                    // Assert access by index
                    Assert.That(setup, Is.EqualTo(collection[indexExpected]));
                });
            }
            Console.WriteLine($"Elapsed time registering: {_stopwatch.Elapsed.TotalMilliseconds:F3}ms");
            // Assert collections are equivalent
            Assert.That(externalPlainCollection, Has.Count.EqualTo(collection.Count));
            Assert.That(externalPlainCollection, Is.EquivalentTo(collection.GetRegisteredItems().ToList()));
        }

        [Test, Category("StyleEquality"), Category("ISetupCollection")]
        [TestCaseSource(nameof(Data_IndexedSetupCollection_NumFormat))]
        public void SetupCollectionNumFormat_Register_MatchEquivalent(Dictionary<int, string> builtIn, List<string> nonBuiltIn)
        {
            ISetupCollection<NumberingFormatSetup> collection = new NumberingFormatSetupCollection();
            foreach (var format in builtIn)
            {
                var setup = new NumberingFormatSetup(format.Value);

                _stopwatch.Start();
                int resultIndex = collection.Register(setup);
                _stopwatch.Stop();

                Assert.Multiple(() =>
                {
                    Assert.That(resultIndex, Is.EqualTo(format.Key));
                    Assert.That(setup.Index, Is.EqualTo(format.Key));
                    Assert.That(setup.Build(), Is.Not.Null);
                    Assert.That(setup, Is.EqualTo(collection[resultIndex]));
                });
                // Ensure that GetRegisteredItems does not return builtin
                var items = collection.GetRegisteredItems();
                Assert.That(items, Does.Not.Contain(setup));
            }
            List<NumberingFormatSetup> externalPlainCollection = new List<NumberingFormatSetup>();
            foreach (var format in nonBuiltIn)
            {
                var setup = new NumberingFormatSetup(format);
                if (!externalPlainCollection.Contains(setup))
                    externalPlainCollection.Add(setup);
                int indexExpected = externalPlainCollection.IndexOf(setup);

                _stopwatch.Start();
                int resultIndex = collection.Register(setup);
                _stopwatch.Stop();

                Assert.Multiple(() =>
                {
                    // Assert indexing on the right position
                    Assert.That(resultIndex, Is.EqualTo(indexExpected + NumberingFormatSetup.StartIndexNotBuiltin));
                    Assert.That(setup.Index, Is.EqualTo(indexExpected + NumberingFormatSetup.StartIndexNotBuiltin));
                    Assert.That(setup.Build(), Is.Not.Null);
                    // Assert access by index
                    Assert.That(setup, Is.EqualTo(collection[indexExpected + NumberingFormatSetup.StartIndexNotBuiltin]));
                });
            }
            Console.WriteLine($"Elapsed time registering: {_stopwatch.Elapsed.TotalMilliseconds:F3}ms");
            // Assert collections are equivalent
            Assert.That(externalPlainCollection, Has.Count.EqualTo(collection.Count));
            Assert.That(externalPlainCollection, Is.EquivalentTo(collection.GetRegisteredItems().ToList()));
        }

        [Test, Category("StyleEquality"), Category("ISetupCollection")]
        public void SetupCollection_BuildContainer_AsExpected()
        {
            // Create Each collection
            ISetupCollection<FillSetup> fillCollection = new IndexedSetupCollection<FillSetup>();
            ISetupCollection<FontSetup> fontCollection = new IndexedSetupCollection<FontSetup>();
            ISetupCollection<BorderSetup> borderCollection = new IndexedSetupCollection<BorderSetup>();
            ISetupCollection<FormatSetup> formatCollection = new IndexedSetupCollection<FormatSetup>();
            ISetupCollection<SharedStringSetup> stringCollection = new SharedStringSetupCollection();
            ISetupCollection<NumberingFormatSetup> numFormatCollection = new NumberingFormatSetupCollection();

            // Retrieve testData
            List<FormatSetup> formatSetups = (List<FormatSetup>)Data_IndexedSetupCollection_FormatSetup.ToList().First()[0];
            List<SharedStringSetup> sharedStrings = (List<SharedStringSetup>)Data_SharedStringSetupCollection.ToList().First()[0];
            Dictionary<int, string> numberingFormatsBuiltIn = (Dictionary<int, string>)Data_IndexedSetupCollection_NumFormat.ToList().First()[0];
            List<string> numberingFormatsCustom = (List<string>)Data_IndexedSetupCollection_NumFormat.ToList().First()[1];
            List<NumberingFormatSetup> numberingSetups = new List<NumberingFormatSetup>();
            numberingSetups.AddRange(numberingFormatsBuiltIn.Values.Select(n => new NumberingFormatSetup(n)));
            numberingSetups.AddRange(numberingFormatsCustom.Select(n => new NumberingFormatSetup(n)));

            // Fill Each collection
            formatSetups.Select(f => f.Border).Where(s => s != null).ToList().ForEach(f => borderCollection.Register(f!));
            formatSetups.Select(f => f.Fill).Where(s => s != null).ToList().ForEach(f => fillCollection.Register(f!));
            formatSetups.Select(f => f.Font).Where(s => s != null).ToList().ForEach(f => fontCollection.Register(f!));
            numberingSetups.ForEach(n => numFormatCollection.Register(n));
            formatSetups.ForEach(f => formatCollection.Register(f!));
            sharedStrings.ForEach(s => stringCollection.Register(s!));

            // Build each container for each collection
            OpenXmlElement borderContainer = borderCollection.BuildContainer<DocumentFormat.OpenXml.Spreadsheet.Borders>();
            OpenXmlElement fillContainer = fillCollection.BuildContainer<DocumentFormat.OpenXml.Spreadsheet.Fills>();
            OpenXmlElement fontContainer = fontCollection.BuildContainer<DocumentFormat.OpenXml.Spreadsheet.Fonts>();
            OpenXmlElement formatContainer = formatCollection.BuildContainer<DocumentFormat.OpenXml.Spreadsheet.Formats>();
            OpenXmlElement stringContainer = stringCollection.BuildContainer<DocumentFormat.OpenXml.Spreadsheet.SharedStringTable>();
            OpenXmlElement numFormatContainer = numFormatCollection.BuildContainer<DocumentFormat.OpenXml.Spreadsheet.NumberingFormats>();

            // Assert that container has the structure expected for each collection
            AssertSetupCollectionBuildContainer(borderContainer, borderCollection);
            AssertSetupCollectionBuildContainer(fillContainer, fillCollection);
            AssertSetupCollectionBuildContainer(fontContainer, fontCollection);
            AssertSetupCollectionBuildContainer(formatContainer, formatCollection);
            AssertSetupCollectionBuildContainer(stringContainer, stringCollection);
            AssertNumFormatCollectionBuildContainer(numFormatContainer, numFormatCollection);
        }

        #region Assert BuildContainer
        private static void AssertSetupCollectionBuildContainer<TSetup>(OpenXmlElement container, ISetupCollection<TSetup> collection)
            where TSetup : Setup, IEquatable<TSetup>, IIndexedSetup
        {
            Assert.That(container.ChildElements.Count, Is.EqualTo(collection.Count));
            for (int i = 0; i < container.ChildElements.Count; i++)
            {
                OpenXmlElement child = container.ChildElements[i];
                TSetup? collectionSetup = collection[i];
                Assert.That(collectionSetup, Is.Not.Null);
                OpenXmlElement setupBuilt = collectionSetup.Build();
                // Compare it is the same using string comparator, not by reference because OpenXmlElement does not override IEquatable
                Assert.That(child.InnerXml, Is.EqualTo(setupBuilt.InnerXml));
            }
        }

        private static void AssertNumFormatCollectionBuildContainer(OpenXmlElement container, ISetupCollection<NumberingFormatSetup> collection)
        {
            Assert.That(container.ChildElements.Count, Is.EqualTo(collection.Count));
            for (int i = 0; i < container.ChildElements.Count; i++)
            {
                OpenXmlElement child = container.ChildElements[i];
                NumberingFormatSetup? collectionSetup = collection[i + NumberingFormatSetup.StartIndexNotBuiltin];
                Assert.That(collectionSetup, Is.Not.Null);
                OpenXmlElement setupBuilt = collectionSetup.Build();
                // Compare it is the same using string comparator, not by reference because OpenXmlElement does not override IEquatable
                Assert.That(child.InnerXml, Is.EqualTo(setupBuilt.InnerXml));
            }
        }
        #endregion

        #region TestCaseSource CombineStyles
        private static IEnumerable<object[]> Data_CombineStylesCases
        {
            get
            {
                var style1 = new Style();
                var style2 = new Style();
                var styleR = new Style();
                style1.Font = new();
                style2.Font = new();
                styleR.Font = new();

                style1.Font.Size = 0;
                style2.Font.Size = 14;
                styleR.Font.Size = 0;
                yield return new object[] { style1, style2, styleR };
                style1.Font.Size = null;
                style2.Font.Size = 14;
                styleR.Font.Size = 14;
                yield return new object[] { style1, style2, styleR };
                style1.Font.Color = null;
                style2.Font.Color = System.Drawing.Color.AliceBlue;
                styleR.Font.Color = System.Drawing.Color.AliceBlue;
                yield return new object[] { style1, style2, styleR };

                style1 = new Style();
                style2 = new Style();
                styleR = new Style();
                style1.Fill = new();
                styleR.Fill = new();
                yield return new object[] { style1, style2, styleR };

                style2.Fill = new();
                style2.Fill.BackgroundColor = System.Drawing.Color.AliceBlue;
                styleR.Fill.BackgroundColor = System.Drawing.Color.AliceBlue;
                yield return new object[] { style1, style2, styleR };
            }

        }
        #endregion

        #region TestCaseSource SetupCollection
        private static IEnumerable<object[]> Data_SharedStringSetupCollection
        {
            get
            {
                ISetupCollection<SharedStringSetup> collection = new SharedStringSetupCollection();
                const string letters = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ";
                Random rnd = new Random();
                List<SharedStringSetup> generated = new List<SharedStringSetup>();
                for (int i = 0; i < _amountIndexedSetupItems; i++)
                {
                    var str = letters.Substring(rnd.Next(letters.Length), 1);
                    generated.Add(new SharedStringSetup(str));
                }
                yield return new object[] { generated, collection };
            }
        }

        private static IEnumerable<object[]> Data_IndexedSetupCollection_Strings
        {
            get
            {
                ISetupCollection<SharedStringSetup> collection = new IndexedSetupCollection<SharedStringSetup>();
                const string letters = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ";
                Random rnd = new Random();
                List<SharedStringSetup> generated = new List<SharedStringSetup>();
                for (int i = 0; i < _amountIndexedSetupItems; i++)
                {
                    var str = letters.Substring(rnd.Next(letters.Length), 1);
                    generated.Add(new SharedStringSetup(str));
                }
                yield return new object[] { generated, collection };
            }
        }

        private static IEnumerable<object[]> Data_IndexedSetupCollection_BorderSetup
        {
            get
            {
                ISetupCollection<BorderSetup> collection = new IndexedSetupCollection<BorderSetup>();
                Random rnd = new Random();
                List<BorderSetup> generated = new List<BorderSetup>();
                int maxBorderTypeIntValue = (int)Enum.GetValues(typeof(BorderStyle.BorderType)).Cast<BorderStyle.BorderType>().Max();
                for (int i = 0; i < _amountIndexedSetupItems; i++)
                {
                    Color randomColor = Color.FromArgb(rnd.Next(256), 0, 0);
                    BorderStyle.BorderType randomBorderType = (BorderStyle.BorderType)rnd.Next(maxBorderTypeIntValue);

                    var style = new BorderStyle();
                    style.Color = randomColor;
                    style.SetBorderType(randomBorderType);
                    var setup = new BorderSetup(style);

                    generated.Add(setup);

                }
                yield return new object[] { generated, collection };
            }
        }

        private static IEnumerable<object[]> Data_IndexedSetupCollection_FillSetup
        {
            get
            {
                ISetupCollection<FillSetup> collection = new IndexedSetupCollection<FillSetup>();
                Random rnd = new Random();
                List<FillSetup> generated = new List<FillSetup>();
                for (int i = 0; i < _amountIndexedSetupItems; i++)
                {
                    Color randomColor = Color.FromArgb(rnd.Next(256), 0, 0);

                    var style = new FillStyle();
                    style.BackgroundColor = randomColor;
                    var setup = new FillSetup(style);

                    generated.Add(setup);

                }
                yield return new object[] { generated, collection };
            }
        }

        private static IEnumerable<object[]> Data_IndexedSetupCollection_FontSetup
        {
            get
            {
                List<string> fonts = new List<string>() { "Calibri", "Times New Roman", "Courier New" };
                ISetupCollection<FontSetup> collection = new IndexedSetupCollection<FontSetup>();
                Random rnd = new Random();
                List<FontSetup> generated = new List<FontSetup>();
                for (int i = 0; i < _amountIndexedSetupItems; i++)
                {
                    Color randomColor = Color.FromArgb(rnd.Next(256), 0, 0);

                    var style = new FontStyle();
                    style.Color = randomColor;
                    style.FontName = fonts[rnd.Next(fonts.Count)];
                    var setup = new FontSetup(style);

                    generated.Add(setup);

                }
                yield return new object[] { generated, collection };
            }
        }

        private static IEnumerable<object[]> Data_IndexedSetupCollection_FormatSetup
        {
            get
            {
                List<FormatSetup> generated = new List<FormatSetup>();
                ISetupCollection<FormatSetup> collection = new IndexedSetupCollection<FormatSetup>();
                // Retrieve the first elements (the only ones) and test with it.
                List<FontSetup> fonts = (List<FontSetup>)Data_IndexedSetupCollection_FontSetup.ToList().First()[0];
                List<FillSetup> fills = (List<FillSetup>)Data_IndexedSetupCollection_FillSetup.ToList().First()[0];
                List<BorderSetup> borders = (List<BorderSetup>)Data_IndexedSetupCollection_BorderSetup.ToList().First()[0];
                for (int i = 0; i < Math.Round(_amountIndexedSetupItems / 2.0); i++) // A half to reduce the test time
                {
                    var setup = new FormatSetup(fills[i], fonts[i], borders[i], null, null);
                    generated.Add(setup);
                }
                yield return new object[] { generated, collection };
            }
        }

        private static IEnumerable<object[]> Data_IndexedSetupCollection_NumFormat
        {
            get
            {
                List<string> generated = new List<string>();
                Dictionary<int, string> builtIn = NumberingFormatSetupCollection.PredefinedFormats;
                const string letters = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ";
                Random rnd = new Random();
                for (int i = 0; i < _amountIndexedSetupItems; i++)
                {
                    var str = letters.Substring(rnd.Next(letters.Length), 1);
                    generated.Add(str);
                }
                yield return new object[] { builtIn, generated };
            }
        }
        #endregion
    }
}
