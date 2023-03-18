using Beporsoft.TabularSheets.Csv;
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

            // Check the ordering sequence
            var orderSequence = Enumerable.Range(0, table.Columns.Count);
            Assert.That(table.Columns.Select(c => c.Order), Is.EquivalentTo(orderSequence));
            Assert.That(table.Columns.Select(c => c.Title).Distinct().Count(), Is.EqualTo(table.Columns.Count));

            foreach (var col in table.Columns)
            {
                Assert.Multiple(() =>
                {
                    // Title not null
                    Assert.That(col.Title, Is.Not.Null);
                    Assert.That(col.Title, Is.Not.EqualTo(string.Empty));
                    Assert.That(col.Owner, Is.EqualTo(table));
                });
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
                Assert.That(originCol2.Title.EndsWith("Col4"), Is.True); // 4th because now is in the 4 position
                Assert.That(originCol3.Title.EndsWith("Col2"), Is.True); // 2th because now is in the 2 position
            });
        }

        /// <summary>
        /// Remove column and preserve order on the following
        /// </summary>
        [Test]
        [TestCase(0)]
        [TestCase(2)]
        [TestCase(5)]
        [TestCase(6)]
        public void RemoveColumn(int positionRemove)
        {
            int x = positionRemove;
            int x1 = x + 1;
            TabularData<Product> table = Generate();

            var oldColX = table.Columns.Single(c => c.Order == x);
            var oldColX1 = table.Columns.SingleOrDefault(c => c.Order == x1);

            table.RemoveColumn(oldColX);

            var newColX = table.Columns.SingleOrDefault(c => c.Order == x);

            Assert.Multiple(() =>
            {
                Assert.That(oldColX1, Is.EqualTo(newColX));
                var orderSequence = Enumerable.Range(0, table.Columns.Count);
                Assert.That(table.Columns.Select(c => c.Order).OrderBy(c => c), Is.EquivalentTo(orderSequence));
            });
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
