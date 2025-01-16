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
        public void ToMarkdown_GenerationOk_Defaults()
        {
            var options = new MarkdownParsingOptions()
            {
            };
            string file = _filesHandler.BuildPath($"{nameof(ToMarkdown_GenerationOk_Defaults)}.md");
            var table = Product.GenerateProductSheet(10);

            table.ToMarkdownTable(file);

            var lines = table.ToMarkdownTable();
        }
        #region Private assert
        private static void AssertMarkdownTable(TabularSheet<Product> table, CsvFixture csvFixture)
        {
            foreach (var col in table.Columns)
            {
                string? name = csvFixture.GetHeaderColumn(col.Index);
                Assert.Multiple(() =>
                {
                    Assert.That(name, Is.Not.Null);
                    Assert.That(name, Is.EqualTo(col.Title));
                });

                for (int i = 0; i < table.Items.Count; i++)
                {
                    Product item = table.Items[i];
                    string cell = csvFixture.GetCell(i, col.Index);
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
