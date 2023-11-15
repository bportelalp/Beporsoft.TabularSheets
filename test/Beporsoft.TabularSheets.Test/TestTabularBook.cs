using Beporsoft.TabularSheets.Test.Helpers;
using Beporsoft.TabularSheets.Test.TestModels;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Test
{
    [Category("Workbook")]
    public class TestTabularBook
    {
        private readonly bool _clearFolderOnEnd = false;
        private readonly int _amountRows = 1000;
        private readonly int _amountReviewsByProduct = 100;
        private readonly TestFilesHandler _filesHandler = new(nameof(TestTabularBook));
        private readonly Stopwatch _stopwatch = new Stopwatch();

        [SetUp]
        public void Setup()
        {
            _stopwatch.Reset();
        }

        [OneTimeTearDown]
        public void OneTimeTearDown()
        {
            if (_clearFolderOnEnd)
                _filesHandler.ClearFiles();

        }


        [Test]
        public void BookAddSheet_ShouldAppendSheetAtEnd()
        {
            const int amountTables = 5;

            TabularBook book = new();
            List<TabularSheet<Product>> tables = new List<TabularSheet<Product>>();
            for (int i = 0; i <= amountTables; i++)
                tables.Add(Product.GenerateProductSheet(_amountRows));

            for (int i = 0; i <= amountTables; i++)
            {
                TabularSheet<Product> sheet = tables[i];
                book.Add(sheet);
                Assert.That(book[i], Is.EqualTo(sheet));
                int index = book.IndexOf(sheet);
                Assert.That(index, Is.EqualTo(i));
            }
        }


        [Test]
        public void BookRemoveSheet_ShouldReturnFalse_IfNotIncluded()
        {
            const int amountTables = 5;

            TabularBook book = new();
            for (int i = 0; i <= amountTables; i++)
                book.Add(Product.GenerateProductSheet(_amountRows));

            TabularSheet<Product> additionalTable = Product.GenerateProductSheet(_amountRows);
            bool removedNotOk = book.Remove(additionalTable);
            Assert.That(removedNotOk, Is.False);
        }


        [Test]
        public void CreateMultiple_ShouldIterateSheetName_IfSameTableName()
        {
            const int amountProducts = 4;
            string path = _filesHandler.BuildPath($"Test{nameof(CreateMultiple_ShouldIterateSheetName_IfSameTableName)}.xlsx");
            List<string> expectedNames = new List<string>();

            TabularBook book = new();
            for (int i = 0; i < amountProducts; i++)
            {
                var productSheet = Product.GenerateProductSheet(_amountRows);
                var productReview = ProductReview.GenerateReviewSheet(productSheet, _amountReviewsByProduct);
                book.Add(productSheet);
                book.Add(productReview);

                expectedNames.Add(i is 0 ? $"{nameof(Product)}" : $"{nameof(Product)}{i}");
                expectedNames.Add(i is 0 ? $"{nameof(ProductReview)}" : $"{nameof(ProductReview)}{i}");
            }

            book.Create(path);
            WorkbookFixture workbook = new(path);
            IEnumerable<string> names = workbook.Sheets.Keys;
            Assert.That(names, Is.EquivalentTo(expectedNames));

        }

        [Test]
        public void CreateMultiple_ShouldIncludeEach_AsExpected()
        {
            TabularBook book = new TabularBook();

            var tableProducts = Product.GenerateProductSheet(_amountRows);
            var tableReviews = ProductReview.GenerateReviewSheet(tableProducts.Items, _amountReviewsByProduct);
            string path = _filesHandler.BuildPath($"Test{nameof(CreateMultiple_ShouldIncludeEach_AsExpected)}.xlsx");

            book.Add(tableProducts);
            book.Add(tableReviews);

            book.Create(path);
            WorkbookFixture workbook = new(path);

            var productFixture = workbook.Sheets[nameof(Product)];
            var reviewFixture = workbook.Sheets[nameof(ProductReview)];
            TabularSheetAsserter.AssertTabularSheet(tableProducts, productFixture);
            TabularSheetAsserter.AssertTabularSheet(tableReviews, reviewFixture);
        }
    }
}
