﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Test.TestModels
{
    /// <summary>
    /// A class which represent a product on a marketstore as test data
    /// </summary>
    public class Product
    {
        public Product()
        {
            Id = Guid.NewGuid();
        }

        public Guid Id { get; set; }
        public string Name { get; set; } = null!;
        public double Cost { get; set; }
        public string Vendor { get; set; } = null!;
        public string CountryOrigin { get; set; } = null!;
        public DateTime LastPriceUpdate { get; set; }
        public TimeSpan DeliveryTime { get; set; }


        /// <summary>
        /// Generate products with fields filled randomly
        /// </summary>
        /// <param name="amount"></param>
        /// <returns></returns>
        internal static IEnumerable<Product> GenerateProducts(int amount = 10)
        {
            const string letters = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZÁÊÌÖ";
            List<Product> products = new();
            foreach (var idx in Enumerable.Range(0, amount))
            {
                var rnd = new Random();
                var product = new Product
                {
                    Name = new string(Enumerable.Repeat(letters, 5).Select(s => s[rnd.Next(s.Length)]).ToArray()).ToLower(),
                    Vendor = new string(Enumerable.Repeat(letters, 10).Select(s => s[rnd.Next(s.Length)]).ToArray()).ToLower(),
                    CountryOrigin = new string(Enumerable.Repeat(letters, 8).Select(s => s[rnd.Next(s.Length)]).ToArray()).ToLower(),
                    Cost = rnd.NextDouble() * 10.0,
                    LastPriceUpdate = new DateTime(2010, 1, 1).AddDays(rnd.Next((DateTime.Now - new DateTime(2010, 1, 1)).Days)),
                    DeliveryTime = TimeSpan.FromSeconds(rnd.NextDouble() * 200000.0)
                };
                products.Add(product);
            }
            return products;
        }

        internal static TabularSheet<Product> GenerateProductSheet(int amount = 1000)
        {
            TabularSheet<Product> table = new();
            table.AddRange(GenerateProducts(amount));

            table.AddColumn(t => t.Id).SetTitle(nameof(Id));
            table.AddColumn(t => t.Name).SetTitle(nameof(Name));
            table.AddColumn(t => t.Vendor).SetTitle(nameof(Vendor));
            table.AddColumn(t => t.CountryOrigin).SetTitle(nameof(CountryOrigin));
            table.AddColumn(t => t.Cost).SetTitle(nameof(Cost));
            table.AddColumn(t => t.LastPriceUpdate).SetTitle(nameof(LastPriceUpdate));
            table.AddColumn(t => t.DeliveryTime).SetTitle(nameof(DeliveryTime));
            return table;
        }
    }
}
