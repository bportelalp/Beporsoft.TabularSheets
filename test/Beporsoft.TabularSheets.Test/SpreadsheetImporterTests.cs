using Beporsoft.TabularSheets.Builders;
using Beporsoft.TabularSheets.Test.Helpers;
using Beporsoft.TabularSheets.Test.TestModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Test
{
    [Category("SpreadsheetImporter")]
    public class SpreadsheetImporterTests
    {
        private readonly TestFilesHandler _filesHandler = new TestFilesHandler("SpreadsheetImport");

        [Test]
        public void CreateSheet_ThenImport_SameData()
        {
            var table = Product.GenerateProductSheet(100);

            string path = _filesHandler.BuildPath($"Test{nameof(CreateSheet_ThenImport_SameData)}.xlsx");
            table.Create(path);

            TabularSheet<Product> import = CopyTableStructure(table);
            SpreadsheetImporter<Product> importer = new(import);
            importer.ImportContent(path);
            AssertProductEqual(table, import);
        }

        [Test]
        public void CreateSheet_ThrowsException_IfPropertyReadonly()
        {
            
            var original = Product.GenerateProductSheet(10);

            string path = _filesHandler.BuildPath($"Test{nameof(CreateSheet_ThrowsException_IfPropertyReadonly)}.xlsx");
            original.Create(path);

            TabularSheet<ProductReadonly> import = [];
            import.Title = nameof(Product);
            // These two properties are
            import.AddColumn(nameof(Product.Id), p => p.Id);
            import.AddColumn(nameof(Product.Name), p => p.Name);
            SpreadsheetImporter<ProductReadonly> importer = new(import);

            Assert.Throws<SheetImportException>(() => importer.ImportContent(path));
        }

        [Test]
        public void CreateSheet_ThrowsException_IfPropertyComplexExpression()
        {

            var original = Product.GenerateProductSheet(10);

            string path = _filesHandler.BuildPath($"Test{nameof(CreateSheet_ThrowsException_IfPropertyReadonly)}.xlsx");
            original.Create(path);

            TabularSheet<Product> import = [];
            import.Title = nameof(Product);
            // These two properties are
            import.AddColumn(nameof(Product.Id), p => p.Id);
            import.AddColumn(nameof(Product.Name), p => $"{p.Name}-Something");
            import.AddColumn(nameof(Product.Cost), p => p.Cost * 2);
            SpreadsheetImporter<Product> importer = new(import);

            Assert.Throws<SheetImportException>(() => importer.ImportContent(path));
        }


        private void AssertProductEqual(TabularSheet<Product> originalTable, TabularSheet<Product> importedTable)
        {
            for(int i = 0; i < originalTable.Count; i++)
            {
                Product original = originalTable[i];
                Product imported = importedTable[i];
                Assert.That(imported.Id, Is.EqualTo(original.Id));
                Assert.That(imported.Name, Is.EqualTo(original.Name));
                Assert.That(imported.CountryOrigin, Is.EqualTo(original.CountryOrigin));
                Assert.That(imported.Cost, Is.EqualTo(original.Cost));
                Assert.That(imported.LastPriceUpdate, Is.EqualTo(original.LastPriceUpdate));
                Assert.That(Math.Abs(imported.DeliveryTime.TotalMilliseconds - original.DeliveryTime.TotalMilliseconds), 
                    Is.LessThan(TimeSpan.FromMilliseconds(1).TotalMilliseconds));
            }
        }

        private TabularSheet<T> CopyTableStructure<T>(TabularSheet<T> original)
        {
            TabularSheet<T> import = new();
            foreach (var col in original.Columns)
            {
                import.AddColumn(col.Title, col.CellContent);
            }
            return import;
        }

        private class ProductReadonly
        {
            public Guid Id { get; }
            public string Name { get; } = string.Empty;
        }
    }
}
