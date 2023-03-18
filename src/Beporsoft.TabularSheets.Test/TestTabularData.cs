using Beporsoft.TabularSheet.Csv;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Test
{
    internal class TestTabularData
    {

        /// <summary>
        /// Create a table and verify correct ordering.
        /// </summary>
        [Test]
        public void TableCreation()
        {
            TabularData<Product> table = Generate();

            TabularDataColumn<Product>? lastCol = null;
            foreach (var col in table.Columns)
            {
                if (lastCol is not null)
                    Assert.That(col.Order, Is.EqualTo(lastCol.Order + 1));
                Assert.That(col, Is.Not.Null);
                Assert.That(col.Owner, Is.EqualTo(table));
            }
        }

        /// <summary>
        /// Reorder a column and verify the names and the reordering of the rest of columns
        /// </summary>
        [Test]
        public void TableReorderColumns()
        {
            TabularData<Product> table = Generate();

            var originCol2 = table.Columns.Skip(2).First();
            var originCol3 = table.Columns.Skip(3).First();
            var originCol4 = table.Columns.Skip(4).First();
            Assert.Multiple(() =>
            {
                Assert.That(originCol2.Order, Is.EqualTo(2));
                Assert.That(originCol4.Order, Is.EqualTo(4));
            });
            originCol2.SetPosition(4);
            Assert.Multiple(() =>
            {
                Assert.That(originCol2.Order, Is.EqualTo(4)); // The set position is correct
                Assert.That(table.Columns.Single(c => c.Order == 4), Is.EqualTo(originCol2)); // The col in 4th position is the original col2
                // There are columns in col 2, 3 and 4
                Assert.That(table.Columns.Any(c => c.Order == 4), Is.True);
                Assert.That(table.Columns.Any(c => c.Order == 3), Is.True);
                Assert.That(table.Columns.Any(c => c.Order == 2), Is.True);

                // The column in position 3 is the column which was previously in position 4
                Assert.That(table.Columns.Single(c => c.Order == 3), Is.EqualTo(originCol4));
                // The column in position 2 is the column which was previously in position 3
                Assert.That(table.Columns.Single(c => c.Order == 2), Is.EqualTo(originCol3));

                // There aren't duplicities of order
                Assert.That(table.Columns.GroupBy(c => c.Order).Any(group => group.Count() > 1), Is.False);

                // The names, which weren't previously setted, keep coherence.
                Assert.That(originCol2.Title, Is.EqualTo("Column4")); // 4th because now is in the 4 position
                Assert.That(originCol3.Title, Is.EqualTo("Column2")); // 2th because now is in the 2 position
            });
        }

        [Test]
        public void RemoveColumn()
        {
            TabularData<Product> table = Generate();

            var col3 = table.Columns.Single(c => c.Order == 3);
            var col4 = table.Columns.Single(c => c.Order == 4);
            table.RemoveColumn(col3);

            var newCol3 = table.Columns.Single(c => c.Order == 3);
            Assert.That(table.Columns.Contains(col3), Is.False);
            Assert.That(newCol3, Is.EqualTo(col4));
            Assert.That(col4.Order, Is.EqualTo(3));
        }

        private static TabularData<Product> Generate()
        {
            TabularSpreadsheet<Product> table = new();
            table.AddRange(Product.GenerateProducts(50));

            TabularDataColumn<Product> col;
            col = table.AddColumn(t => t.Id);
            col = table.AddColumn(t => t.Name);
            col = table.AddColumn(t => t.Vendor);
            col = table.AddColumn(t => t.CountryOrigin);
            col = table.AddColumn("Cost by unit", t => t.Cost);
            col = table.AddColumn(t => t.LastPriceUpdate).SetTitle("Price updated on");
            col = table.AddColumn(t => t.LastUpdate);
            return table;
        }

    }
}
