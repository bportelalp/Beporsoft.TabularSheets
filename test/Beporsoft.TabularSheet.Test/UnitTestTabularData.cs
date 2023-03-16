namespace Beporsoft.TabularSheet.Test
{
    public class Tests
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void Create()
        {
            TabularData<Product> table = new TabularData<Product>();
            table.AddRange(GenerateProducts());
            table.SetColumn(t => t.Id);
            table.SetColumn(t => t.Name);
            table.SetColumn(t => t.Cost);
            table.SetColumn(t => t.LastPriceUpdate);
        }




        private IEnumerable<Product> GenerateProducts()
        {
            yield return new Product("Bread", 0.75, new DateTime(2023,1,1));
            yield return new Product("Milk", 1.10, new DateTime(2023,1,15));
            yield return new Product("Oranges", 0.75, new DateTime(2023,2,1));
        }


        private class Product
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
        }
    }
}