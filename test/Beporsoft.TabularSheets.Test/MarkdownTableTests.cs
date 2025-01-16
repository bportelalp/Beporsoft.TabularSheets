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

        [Test]
        public void ToMarkdownLines_GenerationOk_Defaults()
        {
            string file = _filesHandler.BuildPath($"{nameof(ToMarkdownLines_GenerationOk_Defaults)}.md");
            var table = Product.GenerateProductSheet(10);

            var mdLines = table.ToMarkdownTable();
            var mdFixture = new MarkdownFixture(mdLines, null);
            AssertMarkdownTable(table, mdFixture);
        }

        [Test]
        public void ToMarkdownFile_GenerationOk_Defaults()
        {
            string file = _filesHandler.BuildPath($"{nameof(ToMarkdownFile_GenerationOk_Defaults)}.md");
            var table = Product.GenerateProductSheet(10);

            table.ToMarkdownTable(file);
            var mdFixture = new MarkdownFixture(file, null);
            AssertMarkdownTable(table, mdFixture);
        }

        [Test]
        public void ToMarkdownStream_GenerationOk_Defaults()
        {
            string file = _filesHandler.BuildPath($"{nameof(ToMarkdownStream_GenerationOk_Defaults)}.md");
            var table = Product.GenerateProductSheet(10);

            using (var fs = new FileStream(file, FileMode.Create, FileAccess.Write))
            {
                table.ToMarkdownTable(fs);
            }
            var mdFixture = new MarkdownFixture(file, null);
            AssertMarkdownTable(table, mdFixture);
        }

        [Test]
        public void ToMarkdownLines_SkipTitles_GeneratesHeadersWithoutTitles()
        {
            var options = new MarkdownTableOptions()
            {
                SupressHeaderTitles = true
            };
            string file = _filesHandler.BuildPath($"{nameof(ToMarkdownLines_SkipTitles_GeneratesHeadersWithoutTitles)}.md");
            var table = Product.GenerateProductSheet(10);

            table.ToMarkdownTable(file, options);
            var mdFixture = new MarkdownFixture(file, options);
            AssertMarkdownTable(table, mdFixture);
        }

        #region Private assert
        private static void AssertMarkdownTable(TabularSheet<Product> table, MarkdownFixture mdFixture)
        {
            foreach (var col in table.Columns)
            {
                string? name = mdFixture.GetHeaderColumn(col.Index);
                Assert.Multiple(() =>
                {
                    Assert.That(name, Is.Not.Null);
                    string expectedTitle = mdFixture.Options.SupressHeaderTitles ? string.Empty : col.Title;
                        Assert.That(name, Is.EqualTo(expectedTitle));
                });

                for (int i = 0; i < table.Items.Count; i++)
                {
                    Product item = table.Items[i];
                    string cell = mdFixture.GetCell(i, col.Index);
                    Assert.That(cell, Is.Not.Null);
                    object value = col.Apply(item);
                    Assert.That(cell, Is.EqualTo(value.ToString()));
                }
            }
        }
        #endregion

        #region Data
        private static IEnumerable<object[]> DataOptions_ToCsv_GenerationOk
        {
            get
            {
                // English, comma separator usually ., so csv must use semicolon or comma
                yield return new object[] { CultureInfo.GetCultureInfo("en-US"), CsvOptions.SemicolonSeparator, Encoding.UTF8 };
                yield return new object[] { CultureInfo.GetCultureInfo("en-US"), CsvOptions.CommaSeparator, Encoding.UTF8 };
                yield return new object[] { CultureInfo.GetCultureInfo("en-US"), CsvOptions.CommaSeparator, Encoding.GetEncoding("latin1") };
                // Spanish, comma separator usually , so csv must use only semicolon
                yield return new object[] { CultureInfo.GetCultureInfo("es-ES"), CsvOptions.SemicolonSeparator, Encoding.UTF8 };
                yield return new object[] { CultureInfo.GetCultureInfo("es-ES"), CsvOptions.SemicolonSeparator, Encoding.GetEncoding("latin1") };
                yield return new object[] { CultureInfo.GetCultureInfo("es-ES"), CsvOptions.SemicolonSeparator, Encoding.UTF32 };
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
