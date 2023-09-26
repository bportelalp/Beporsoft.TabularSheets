using Beporsoft.TabularSheets.Builders.Interfaces;
using Beporsoft.TabularSheets.Builders.SheetBuilders;
using Beporsoft.TabularSheets.Builders.StyleBuilders;
using Beporsoft.TabularSheets.Builders.StyleBuilders.SetupCollections;
using Beporsoft.TabularSheets.CellStyling;
using System.Diagnostics;
using System.Drawing;

namespace Beporsoft.TabularSheets.Test
{
    [Category("Stylesheet")]
    internal class TestCellStyling
    {

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
        [TestCaseSource(nameof(CombineStylesCases))]
        public void StyleCombiner_Combine_IEquatableCompliant(Style style1, Style style2, Style styleExpected)
        {
            Style styleResult = StyleCombiner.Combine(style1, style2);
            Assert.That(styleResult, Is.EqualTo(styleExpected));
        }

        private static IEnumerable<object[]> CombineStylesCases()
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


        [Test, Category("StyleEquality")]
        [TestCaseSource(nameof(Data_SharedStringSetupCollection))]
        [TestCaseSource(nameof(Data_IndexedSetupCollection_BorderSetup))]
        [TestCaseSource(nameof(Data_IndexedSetupCollection_FillSetup))]
        public void SetupCollection_Register_MatchEquivalentWhenItIs<TSetup>(List<TSetup> randomSetups, ISetupCollection<TSetup> collection)
            where TSetup : Setup, IEquatable<TSetup>, IIndexedSetup
        {
            Stopwatch sw  = new Stopwatch();
            List<TSetup> externalPlainCollection = new List<TSetup>();
            foreach (var setup in randomSetups)
            {
                if (!externalPlainCollection.Contains(setup))
                    externalPlainCollection.Add(setup);
                int indexExpected = externalPlainCollection.IndexOf(setup);

                sw.Start();
                int resultIndex = collection.Register(setup);
                sw.Stop();

                Assert.Multiple(() =>
                {
                    Assert.That(resultIndex, Is.EqualTo(indexExpected));
                    Assert.That(setup.Index, Is.EqualTo(indexExpected));
                });
            }
            Assert.That(externalPlainCollection, Has.Count.EqualTo(collection.Count));
            Assert.That(externalPlainCollection, Is.EquivalentTo(collection.GetRegisteredItems().ToList()));
            Console.WriteLine($"Elapsed time registering: {sw.Elapsed.TotalMilliseconds:F3}ms");
        }

        private static IEnumerable<object[]> Data_SharedStringSetupCollection
        {
            get
            {
                ISetupCollection<SharedStringSetup> collection = new SharedStringSetupCollection();
                const string letters = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ";
                Random rnd = new Random();
                List<SharedStringSetup> generated = new List<SharedStringSetup>();
                for (int i = 0; i < 10000; i++)
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
                for (int i = 0; i < 10000; i++)
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
                for (int i = 0; i < 10000; i++)
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

    }
}
