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
        public void ImportSheet_Ok()
        {
            var table = Product.GenerateProductSheet(100);

            string path = _filesHandler.BuildPath($"Test{nameof(ImportSheet_Ok)}.xlsx");
            table.Create(path);

            TabularSheet<Product> import = new();
            foreach (var col in table.Columns)
            {
                import.AddColumn(col.Title, col.CellContent);
            }
            SpreadsheetImporter<Product> importer = new(import);
            importer.ImportContent(path);
            AssertImport(table, import);
            

        }

        public void AssertImport(TabularSheet<Product> original, TabularSheet<Product> imported)
        {
            for(int i = 0; i <= original.Count; i++)
            {
                Assert.That(imported[i], Is.EqualTo(original[i]));
            }
        }
    }
}
