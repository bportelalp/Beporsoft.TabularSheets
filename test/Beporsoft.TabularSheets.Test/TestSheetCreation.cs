using Beporsoft.TabularSheets.Test.Helpers;
using Beporsoft.TabularSheets.Test.TestModels;
using Beporsoft.TabularSheets.Tools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Test
{

    public class TestSheetCreation
    {
        private readonly TestFilesHandler _filesHandler = new TestFilesHandler("SheetCreation");

        [Test, Category("StreamCreation")]
        [TestCaseSource(nameof(Data_FileHelpers_Verifypath))]
        public void FileHelpers_Verifypath_Ok(string? correctedPath, string providedPath, string[] targetExtensions)
        {
            string verify() => FileHelpers.VerifyPath(providedPath, targetExtensions);
            if (correctedPath is null)
            {
                // Must throw exception
                Assert.Throws<FileLoadException>(() => verify());
            }
            else
            {
                // Must give a result
                Assert.That(verify(), Is.EqualTo(correctedPath));
            }
        }


        [Test, Category("StreamCreation")]
        public void Create_CheckFileName()
        {
            TabularSheet<Product> table = Product.GenerateProductSheet();
            string pathOk = _filesHandler.BuildPath($"{nameof(Create_CheckFileName)}.xlsx");
            string pathNotOk = _filesHandler.BuildPath($"{nameof(Create_CheckFileName)}.xls");
            string pathWrongExtension = _filesHandler.BuildPath($"{nameof(Create_CheckFileName)}.csv");
            Assert.Multiple(() =>
            {
                Assert.That(() => table.Create(pathOk), Throws.Nothing);
                Assert.Catch<FileLoadException>(() => table.Create(pathNotOk));
                Assert.Catch<FileLoadException>(() => table.Create(pathWrongExtension));
            });
        }

        [Test, Category("StreamCreation")]
        public void Create_Ok_AsMemoryStream()
        {
            TabularSheet<Product> table = Product.GenerateProductSheet();

            using (var ms = table.Create())
            {
                var sheetFixture = new SheetFixture(ms);
                TabularSheetAsserter.AssertTabularSheet(table, sheetFixture);
            }
        }

        [Test, Category("StreamCreation")]
        public void Create_Ok_OnStreamWhichIsMemoryStream()
        {
            TabularSheet<Product> table = Product.GenerateProductSheet();

            using (var ms = new MemoryStream())
            {
                table.Create(ms);
                var sheetFixture = new SheetFixture(ms);
                TabularSheetAsserter.AssertTabularSheet(table, sheetFixture);
            }
        }

        [Test, Category("StreamCreation")]
        public void Create_Ok_OnStreamWhichIsFileStream()
        {
            TabularSheet<Product> table = Product.GenerateProductSheet();
            string path = _filesHandler.BuildPath($"Test{nameof(Create_Ok_OnStreamWhichIsFileStream)}.xlsx");

            using (var fs = new FileStream(path, FileMode.Create, FileAccess.ReadWrite))
            {
                table.Create(fs);
            }

            var sheetFixture = new SheetFixture(path);
            TabularSheetAsserter.AssertTabularSheet(table, sheetFixture);
        }

        [Test, Category("StreamCreation")]
        public void CreateWorkbook_Ok_AsMemoryStream()
        {
            TabularBook book = new TabularBook();

            var tableProducts = Product.GenerateProductSheet(100);
            var tableReviews = ProductReview.GenerateReviewSheet(tableProducts.Items, 10);

            book.Add(tableProducts);
            book.Add(tableReviews);

            using MemoryStream ms = book.Create();
            WorkbookFixture workbook = new(ms);

            var productFixture = workbook.Sheets[nameof(Product)];
            var reviewFixture = workbook.Sheets[nameof(ProductReview)];
            TabularSheetAsserter.AssertTabularSheet(tableProducts, productFixture);
            TabularSheetAsserter.AssertTabularSheet(tableReviews, reviewFixture);
        }

        [Test, Category("StreamCreation")]
        public void CreateWorkbook_Ok_OnStreamWhichIsMemoryStream()
        {
            TabularBook book = new TabularBook();

            var tableProducts = Product.GenerateProductSheet(100);
            var tableReviews = ProductReview.GenerateReviewSheet(tableProducts.Items, 10);
            book.Add(tableProducts);
            book.Add(tableReviews);

            using (var ms = new MemoryStream())
            {
                book.Create(ms);
                WorkbookFixture workbook = new(ms);

                var productFixture = workbook.Sheets[nameof(Product)];
                var reviewFixture = workbook.Sheets[nameof(ProductReview)];
                TabularSheetAsserter.AssertTabularSheet(tableProducts, productFixture);
                TabularSheetAsserter.AssertTabularSheet(tableReviews, reviewFixture);
            }
        }

        [Test, Category("StreamCreation")]
        public void CreateWorkbook_Ok_OnStreamWhichIsFileStream()
        {
            string path = _filesHandler.BuildPath($"Test{nameof(CreateWorkbook_Ok_OnStreamWhichIsFileStream)}.xlsx");

            TabularBook book = new TabularBook();

            var tableProducts = Product.GenerateProductSheet(100);
            var tableReviews = ProductReview.GenerateReviewSheet(tableProducts.Items, 10);
            book.Add(tableProducts);
            book.Add(tableReviews);

            using (var fs = new FileStream(path, FileMode.Create, FileAccess.ReadWrite))
            {
                book.Create(fs);
            }

            WorkbookFixture workbook = new(path);

            var productFixture = workbook.Sheets[nameof(Product)];
            var reviewFixture = workbook.Sheets[nameof(ProductReview)];
            TabularSheetAsserter.AssertTabularSheet(tableProducts, productFixture);
            TabularSheetAsserter.AssertTabularSheet(tableReviews, reviewFixture);
        }


        [OneTimeTearDown]
        public void OneTimeTearDown()
        {
            _filesHandler.ClearFiles();
        }

        private static IEnumerable<object?[]> Data_FileHelpers_Verifypath
        {
            get
            {
                yield return new object?[] { "File.xlsx", "File.xlsx", new string[] { ".xlsx", ".xls" } };
                yield return new object?[] { "File.xlsx", "File.xlsx", SpreadsheetFileExtension.AllowedExtensions };
                yield return new object?[] { "File.xls", "File.xls", new string[] { ".xlsx", ".xls" } };
                yield return new object?[] { "File.xlsx", "File", new string[] { ".xlsx", ".xls" } };
                yield return new object?[] { "File.xls", "File", new string[] { ".xls", ".xlsx" } };
                yield return new object?[] { "File.csv", "File.csv", new string[] { ".xls", ".csv" } };
                yield return new object?[] { "File.csv", "File.csv", CsvFileExtension.AllowedExtensions };
                yield return new object?[] { "File.xlsx", "File", new string[] { ".xlsx", "csv" } };
                yield return new object?[] { null, "File.csv", new string[] { ".xlsx", ".xls" } };
            }
        }

    }
}
