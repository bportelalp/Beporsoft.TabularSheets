using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheet.Test
{
    internal class Product
    {
        public Product(string name, double cost, DateTime lastUpdate)
        {
            Id = Guid.NewGuid();
            Name = name;
            Cost = cost;
            LastUpdate = lastUpdate;
        }
        public Guid Id { get; set; }
        public string Name { get; set; } = null!;
        public double Cost { get; set; }
        public DateTime LastUpdate { get; }
        public DateTime LastPriceUpdate { get; set; }


        internal static IEnumerable<Product> GenerateProducts()
        {
            yield return new Product("Bread", 0.75, new DateTime(2023, 1, 1));
            yield return new Product("Milk", 1.10, new DateTime(2023, 1, 15));
            yield return new Product("Oranges", 0.75, new DateTime(2023, 2, 1));
        }
    }
}
