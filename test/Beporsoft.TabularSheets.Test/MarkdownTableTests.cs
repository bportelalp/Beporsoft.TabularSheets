using Beporsoft.TabularSheets.Options;
using Beporsoft.TabularSheets.Test.Helpers;
using Beporsoft.TabularSheets.Test.TestModels;
using System.Globalization;
using System.Text;

namespace Beporsoft.TabularSheets.Test
{
    [Category("MarkdownTable")]
    internal class MarkdownTableTests
    {
        private readonly bool _clearFolderOnEnd = false;
        private readonly TestFilesHandler _filesHandler = new(nameof(MarkdownTableTests));
        private readonly MarkdownTableOptions _basicOpt = new MarkdownTableOptions()
        {
            CompactMode = true
        };

        [Test]
        public void ToMarkdownLines_GenerationOk_Defaults()
        {
            string file = _filesHandler.BuildPath($"{nameof(ToMarkdownLines_GenerationOk_Defaults)}.md");
            var table = Product.GenerateProductSheet(10);

            var mdLines = table.ToMarkdownTable(_basicOpt);
            var mdFixture = new MarkdownFixture(mdLines, _basicOpt);
            AssertMarkdownTable(table, mdFixture, false);
        }

        [Test]
        public void ToMarkdownFile_GenerationOk_Defaults()
        {
            string file = _filesHandler.BuildPath($"{nameof(ToMarkdownFile_GenerationOk_Defaults)}.md");
            var table = Product.GenerateProductSheet(10);

            table.ToMarkdownTable(file, _basicOpt);
            var mdFixture = new MarkdownFixture(file, _basicOpt);
            AssertMarkdownTable(table, mdFixture);
        }

        [Test]
        public void ToMarkdownStream_GenerationOk_Defaults()
        {
            string file = _filesHandler.BuildPath($"{nameof(ToMarkdownStream_GenerationOk_Defaults)}.md");
            var table = Product.GenerateProductSheet(10);

            using (var fs = new FileStream(file, FileMode.Create, FileAccess.Write))
            {
                table.ToMarkdownTable(fs, _basicOpt);
            }
            var mdFixture = new MarkdownFixture(file, _basicOpt);
            AssertMarkdownTable(table, mdFixture);
        }

        [Test]
        public void ToMarkdownLines_SkipTitles_GeneratesHeadersWithoutTitles()
        {
            var options = new MarkdownTableOptions()
            {
                SupressHeaderTitles = true,
                CompactMode = true
            };
            string file = _filesHandler.BuildPath($"{nameof(ToMarkdownLines_SkipTitles_GeneratesHeadersWithoutTitles)}.md");
            var table = Product.GenerateProductSheet(10);

            table.ToMarkdownTable(file, options);
            var mdFixture = new MarkdownFixture(file, options);
            AssertMarkdownTable(table, mdFixture);
        }

        [Test]
        public void ToMarkdownLines_ExtendedMode_PadSpacesAlignColumns()
        {
            var options = new MarkdownTableOptions()
            {
                CompactMode = false
            };
            string file = _filesHandler.BuildPath($"{nameof(ToMarkdownLines_ExtendedMode_PadSpacesAlignColumns)}.md");
            var table = Product.GenerateProductSheet(10);

            table.ToMarkdownTable(file, options);
            var mdFixture = new MarkdownFixture(file, options);
            AssertMarkdownTable(table, mdFixture, true);
        }

        #region Private assert
        private static void AssertMarkdownTable(TabularSheet<Product> table, MarkdownFixture mdFixture, bool assertColumnLength = false)
        {
            foreach (var col in table.Columns)
            {
                int colLength = mdFixture.GetHeaderLenght(col.Index);
                string? name = mdFixture.GetHeaderColumn(col.Index);
                Assert.Multiple(() =>
                {
                    Assert.That(name, Is.Not.Null);
                    string expectedTitle = mdFixture.Options.SupressHeaderTitles ? string.Empty : col.Title;
                    Assert.That(name, Is.EqualTo(expectedTitle));
                });

                string? headerSeparator = mdFixture.GetHeaderSeparatorColumn(col.Index);
                Assert.Multiple(() =>
                {
                    Assert.That(headerSeparator, Is.Not.Empty);
                    Assert.That(headerSeparator.ToArray(), Is.All.EqualTo('-')); // All are -
                    Assert.That(headerSeparator.Length, Is.AtLeast(2)); // At least is --
                    if (assertColumnLength)
                        Assert.That(mdFixture.GetHeaderSeparatorLenght(col.Index), Is.EqualTo(colLength));
                });

                for (int i = 0; i < table.Items.Count; i++)
                {
                    Product item = table.Items[i];
                    string cell = mdFixture.GetCell(i, col.Index);
                    Assert.That(cell, Is.Not.Null);
                    object value = col.Apply(item);
                    Assert.That(cell, Is.EqualTo(value.ToString()));
                    if (assertColumnLength)
                        Assert.That(mdFixture.GetCellLength(i, col.Index), Is.EqualTo(colLength));
                }
            }
        }
        #endregion

        [OneTimeTearDown]
        public void OneTimeTearDown()
        {
            if (_clearFolderOnEnd)
                _filesHandler.ClearFiles();
        }
    }
}
