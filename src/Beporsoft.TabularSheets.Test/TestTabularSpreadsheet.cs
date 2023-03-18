using Beporsoft.TabularSheets.Csv;

namespace Beporsoft.TabularSheets.Test
{
    public class TestTabularSpreadsheet
    {

        [Test]
        public void CheckFileName()
        {
            TabularSpreadsheet<Product> table = Generate();
            string pathOk = GetPath("ExcelCheckFileName.xlsx");
            string pathOkAlternative = GetPath("ExcelCheckFileName.xls");
            string pathWrongExtension = GetPath("ExcelCheckFileName.csv");
            Assert.Multiple(() =>
            {
                Assert.That(() => table.Create(pathOk), Throws.Nothing);
                Assert.That(() => table.Create(pathOkAlternative), Throws.Nothing);
                Assert.That(() => table.Create(pathWrongExtension), Throws.Exception);
            });
        }

        private static TabularSpreadsheet<Product> Generate()
        {
            TabularSpreadsheet<Product> table = new();
            table.AddRange(Product.GenerateProducts(50));

            table.AddColumn(t => t.Id).SetTitle("Product Id");
            table.AddColumn(t => t.Name).SetTitle("Product name");
            table.AddColumn(t => t.Vendor).SetTitle("Vendor");
            table.AddColumn(t => t.CountryOrigin).SetTitle("Country");
            table.AddColumn("Cost by unit", t => t.Cost);
            table.AddColumn(t => t.LastPriceUpdate).SetTitle("Price updated on");
            table.AddColumn(t => t.LastUpdate).SetTitle("Last update"); ;
            return table;
        }


        private string GetPath(string fileName)
        {
            DirectoryInfo? projectDir = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory)?.Parent?.Parent?.Parent;
            return $"{projectDir!.FullName}/Results/{fileName}";
        }

    }
}