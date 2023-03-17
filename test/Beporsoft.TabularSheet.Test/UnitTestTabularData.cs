using Beporsoft.TabularSheet.Csv;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheet.Test
{
    internal class UnitTestTabularData
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
            }
        }


        private TabularData<Product> Generate()
        {
            TabularData<Product> table = new TabularData<Product>();
            table.AddRange(Product.GenerateProducts());

            TabularDataColumn<Product> col;
            col = table.SetColumn(t => t.Id);
            col = table.SetColumn(t => t.Name);
            col = table.SetColumn(t => t.Cost);
            col = table.SetColumn(t => t.LastPriceUpdate);
            return table;
        }

    }
}
