using Beporsoft.TabularSheets.CellStyling;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Test.TestsStyles
{
    internal class TestCellStyling
    {

        [Test]
        public void CheckBorderEqualityContract()
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

            border.SetBorderType(BorderStyle.BorderType.None);
            border.Color = null;
            Assert.Multiple(() =>
            {
                Assert.That(borderModified, Is.EqualTo(border));
                Assert.That(borderModified, Is.EqualTo(BorderStyle.Default));
                Assert.That(border, Is.EqualTo(BorderStyle.Default));
            });
        }

        [Test]
        public void CheckFontEqualityContract()
        {
            var font = new FontStyle();
            Assert.That(font, Is.EqualTo(FontStyle.Default));

            font.Size = 12;
            font.Font = "efa";
            font.Color = System.Drawing.Color.Aquamarine;
            Assert.That(font, Is.Not.EqualTo(FontStyle.Default));

            font.Size = null;
            font.Font = null;
            font.Color = null;
            Assert.That(font, Is.EqualTo(FontStyle.Default));
        }

        [Test]
        public void CheckFillEqualityContract()
        {
            var fill = new FillStyle();
            Assert.That(fill, Is.EqualTo(FillStyle.Default));

            fill.BackgroundColor = System.Drawing.Color.Aquamarine;
            Assert.That(fill, Is.Not.EqualTo(FillStyle.Default));

            fill.BackgroundColor = null;
            Assert.That(fill, Is.EqualTo(FillStyle.Default));
        }

        [Test]
        public void CheckStyleEqualityContract()
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

        [Test]
        [TestCaseSource(nameof(CombineStylesCases))]
        public void CombineStyles(Style style1, Style style2, Style styleExpected)
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

    }
}
