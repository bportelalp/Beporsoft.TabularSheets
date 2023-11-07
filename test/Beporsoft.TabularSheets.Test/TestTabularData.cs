﻿using System.Text.RegularExpressions;
using Beporsoft.TabularSheets.Test.TestModels;

namespace Beporsoft.TabularSheets.Test
{
    [Category("SheetStructure")]
    internal class TestTabularData
    {

        /// <summary>
        /// Create a table and verify correct ordering.
        /// </summary>
        [Test, Category("Columns")]
        public void CreateTable_ColumnsAndTitles_AsExpected()
        {
            TabularData<Product> table = Generate();
            // Check the ordering sequence
            var orderSequence = Enumerable.Range(0, table.Columns.Count());
            Assert.Multiple(() =>
            {
                Assert.That(table.Columns.Select(c => c.Index), Is.EquivalentTo(orderSequence));
                Assert.That(table.Columns.Select(c => c.Title).Distinct().Count(), Is.EqualTo(table.Columns.Count()));
            });
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
        [Test, Category("Columns")]
        public void TabularData_ReorderColumns_AllReorganized()
        {
            TabularData<Product> table = Generate();
            var originCol2 = table.Columns.Skip(2).First();
            var originCol3 = table.Columns.Skip(3).First();
            var originCol4 = table.Columns.Skip(4).First();
            Assert.Multiple(() =>
            {
                Assert.That(originCol2.Index, Is.EqualTo(2));
                Assert.That(originCol4.Index, Is.EqualTo(4));
            });
            originCol2.SetIndex(4);
            Assert.Multiple(() =>
            {
                Assert.That(originCol2.Index, Is.EqualTo(4)); // The set position is correct
                Assert.That(table.Columns.Single(c => c.Index == 4), Is.EqualTo(originCol2)); // The col in 4th position is the original col2
                // There are columns in col 2, 3 and 4
                Assert.That(table.Columns.Any(c => c.Index == 4), Is.True);
                Assert.That(table.Columns.Any(c => c.Index == 3), Is.True);
                Assert.That(table.Columns.Any(c => c.Index == 2), Is.True);

                // The column in position 3 is the column which was previously in position 4
                Assert.That(table.Columns.Single(c => c.Index == 3), Is.EqualTo(originCol4));
                // The column in position 2 is the column which was previously in position 3
                Assert.That(table.Columns.Single(c => c.Index == 2), Is.EqualTo(originCol3));

                // There aren't duplicities of order
                Assert.That(table.Columns.GroupBy(c => c.Index).Any(group => group.Count() > 1), Is.False);

                // The names, which weren't previously setted, keep coherence.
                Assert.That(originCol2.Title.EndsWith("Col4"), Is.True); // 4th because now is in the 4 position
                Assert.That(originCol3.Title.EndsWith("Col2"), Is.True); // 2th because now is in the 2 position
            });
        }

        /// <summary>
        /// Ensure that when title of column is empty, the title is the default and when not, is the established
        /// </summary>
        [Test, Category("Columns")]
        public void RenameColumns_Restore_PreserveDefault()
        {
            const string customTitle = "CustomTitle";
            Regex regexDefaultColumnName = new(@"ProductCol\d{0,}");
            TabularData<Product> table = Generate();
            foreach (var column in table.Columns)
            {
                string initialTitle = column.Title;
                if (regexDefaultColumnName.Match(initialTitle).Success)
                {
                    // Make custom title, then delete to restore the initial, which is the autogenerated and include the order
                    column.SetTitle(customTitle);
                    Assert.That(column.Title, Is.EqualTo(customTitle));
                    column.SetTitle(null);
                    Assert.Multiple(() =>
                    {
                        Assert.That(column.Title, Is.EqualTo(initialTitle));
                        Assert.That(column.Title, Does.EndWith(column.Index.ToString()));
                    });
                }
                else
                {
                    // Delete to get a autogenerated title, ensure that match the regex and the order, then undo and chech is the initial
                    column.SetTitle(string.Empty);
                    Assert.Multiple(() =>
                    {
                        Assert.That(regexDefaultColumnName.Match(column.Title).Success, Is.True);
                        Assert.That(column.Title, Does.EndWith(column.Index.ToString()));
                    });
                    column.SetTitle(initialTitle);
                    Assert.That(column.Title, Is.EqualTo(initialTitle));
                }
            }

        }

        /// <summary>
        /// Remove column and preserve order on the following
        /// </summary>
        [Test, Category("Columns")]
        [TestCase(0), TestCase(2), TestCase(5), TestCase(6)]
        public void RemoveColumns_RestoreIndexTheFollowing(int positionRemove)
        {
            int x = positionRemove;
            int x1 = x + 1;
            TabularData<Product> table = Generate();

            var oldColX = table.Columns.Single(c => c.Index == x);
            var oldColX1 = table.Columns.SingleOrDefault(c => c.Index == x1);

            table.RemoveColumn(oldColX);

            var newColX = table.Columns.SingleOrDefault(c => c.Index == x);

            Assert.Multiple(() =>
            {
                Assert.That(oldColX1, Is.EqualTo(newColX));
                var orderSequence = Enumerable.Range(0, table.Columns.Count());
                Assert.That(table.Columns.Select(c => c.Index).OrderBy(c => c), Is.EquivalentTo(orderSequence));
            });
        }

        [Test, Category("Rows")]
        public void Table_AddItems_KeepInvariantOriginalLists()
        {
            IEnumerable<Product> initial = Product.GenerateProducts(10);
            int countInitial = initial.Count();
            TabularSheet<Product> table = new(initial);
            // Items is not the same object as initial
            // But the references contained are the same
            Assert.Multiple(() =>
            {
                Assert.That(initial, Is.Not.SameAs(table.Items));
                Assert.That(initial.SequenceEqual(table.Items), Is.True);
            });
            IEnumerable<Product> addition = Product.GenerateProducts(10);
            table.AddRange(addition);
            Assert.Multiple(() =>
            {
                Assert.That(initial.Count(), Is.EqualTo(countInitial));
                Assert.That(countInitial + addition.Count(), Is.EqualTo(table.Items.Count));
                Assert.That(initial.SequenceEqual(table.Items), Is.False);
                List<Product> outsideMerge = new List<Product>(initial);
                outsideMerge.AddRange(addition);
                Assert.That(table.Items.SequenceEqual(outsideMerge), Is.True);
            });


        }

        private static TabularData<Product> Generate()
        {
            TabularSheet<Product> table = new();
            table.AddRange(Product.GenerateProducts(50));

            TabularDataColumn<Product> col;
            col = table.AddColumn(t => t.Id);
            col = table.AddColumn(t => t.Name);
            col = table.AddColumn(t => t.Vendor);
            col = table.AddColumn(t => t.CountryOrigin);
            col = table.AddColumn("Cost by unit", t => t.Cost);
            col = table.AddColumn(t => t.LastPriceUpdate).SetTitle("Price updated on");
            col = table.AddColumn(t => t.DeliveryTime).SetTitle("Delivery Time");
            return table;
        }

    }
}
