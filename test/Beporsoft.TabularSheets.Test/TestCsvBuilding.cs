using Beporsoft.TabularSheets.Options;
using Beporsoft.TabularSheets.Test.Helpers;
using DocumentFormat.OpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Test
{
    [Category("CSV")]
    internal class TestCsvBuilding
    {
        private readonly bool _clearFolderOnEnd = false;
        private readonly TestFilesHandler _filesHandler = new("Csv");

        [Test]
        [TestCaseSource(nameof(DataOptions_ToCsv_GenerationOk))]
        public void ToCsv_GenerationOk(CultureInfo culture, string separator, Encoding encoding)
        {
            CultureInfo.CurrentCulture = culture;
            var options = new CsvOptions()
            {
                Separator = separator,
                Encoding = encoding
            };
            string file = _filesHandler.BuildPath($"{nameof(ToCsv_GenerationOk)}.csv");
            var table = Product.GenerateTestSheet();
            Assert.Multiple(() =>
            {
                // Test file
                Assert.That(() =>
                {
                    table.ToCsv(file, options);
                    var csvFixture = new CsvFixture(file, options);
                    AssertCsv(table, csvFixture);
                }, Throws.Nothing);
                // Test stream
                Assert.That(() =>
                {
                    using MemoryStream ms = table.ToCsv(options);
                    var csvFixture = new CsvFixture(ms, options);
                    AssertCsv(table, csvFixture);
                }, Throws.Nothing);
            });
        }

        #region Private assert
        private static void AssertCsv(TabularSheet<Product> table, CsvFixture csvFixture)
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
