using DocumentFormat.OpenXml.Office2010.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Test.TestModels
{
    /// <summary>
    /// A class which represent a product review in a marketplace as test data.
    /// </summary>
    public class ProductReview
    {
        public Product Product { get; set; } = default!;
        public string UserName { get; set; } = string.Empty;
        public int Stars { get; set; }
        public int ReviewLikes { get; set; }
        public int ReviewDislikes { get; set; }
        public string ReviewText { get; set; } = string.Empty;


        /// <summary>
        /// Generate products with fields filled randomly
        /// </summary>
        /// <param name="amount"></param>
        /// <returns></returns>
        internal static IEnumerable<IGrouping<Product, ProductReview>> GenerateProductReviews(IEnumerable<Product> products, int maxReviewsByProduct = 10)
        {
            const string letters = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZÁÊÌÖ";
            List<ProductReview> reviews = new List<ProductReview>();
            var rnd = new Random();

            foreach (var product in products)
            {
                int maxReviews = rnd.Next(maxReviewsByProduct);
                foreach (var idx in Enumerable.Range(0, maxReviews))
                {
                    int charsTitle = rnd.Next(100);
                    int charsUsername = rnd.Next(20);

                    var review = new ProductReview
                    {
                        Product = product,
                        UserName = new string(Enumerable.Repeat(letters, charsUsername).Select(s => s[rnd.Next(s.Length)]).ToArray()).ToLower(),
                        ReviewText = new string(Enumerable.Repeat(letters, charsTitle).Select(s => s[rnd.Next(s.Length)]).ToArray()).ToLower(),
                        Stars = rnd.Next(1, 11),
                        ReviewLikes = rnd.Next(1000),
                        ReviewDislikes = rnd.Next(500),
                    };
                    reviews.Add(review);
                }
            }

            return reviews.GroupBy(r => r.Product);
        }

        internal static TabularSheet<IGrouping<Product, ProductReview>> GenerateReviewSheet(IEnumerable<Product> products, int maxReviewsByProduct = 10)
        {
            TabularSheet<IGrouping<Product, ProductReview>> table = new();
            var reviews = GenerateProductReviews(products, maxReviewsByProduct);
            table.AddRange(reviews);

            table.AddColumn(t => t.Key.Name).SetTitle($"{nameof(Product)}{nameof(TestModels.Product.Name)}");
            table.AddColumn(t => t.Average(r => r.Stars)).SetTitle("AverageStars");
            table.AddColumn(t => t.Count()).SetTitle("Total Reviews");
            table.AddColumn(t => t.OrderByDescending(r => r.ReviewLikes).FirstOrDefault()?.ReviewText ?? string.Empty).SetTitle("Most liked review");
            table.AddColumn(t => t.OrderByDescending(r => r.ReviewDislikes).FirstOrDefault()?.ReviewText ?? string.Empty).SetTitle("Worse review");
            table.Title = nameof(ProductReview);
            return table;
        }
    }
}
