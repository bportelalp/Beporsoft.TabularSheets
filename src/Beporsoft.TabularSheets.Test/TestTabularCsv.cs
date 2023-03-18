using Beporsoft.TabularSheets.Csv;

namespace Beporsoft.TabularSheets.Test
{
    public class TestTabularCsv
    {

        [Test]
        public void Creation()
        {
            string path = GetPath("BasicCsv.csv");
            TabularCsv<Product> table = new TabularCsv<Product>();
            table.AddRange(Product.GenerateProducts());
            table.AddColumn(t => t.Id);
            table.AddColumn(t => t.Name);
            table.AddColumn(t => t.Cost);
            table.AddColumn(t => t.LastPriceUpdate);
            table.Create(path);
        }

        [Test]
        public void CreationPathWrong()
        {
            TabularCsv<Product> table = new TabularCsv<Product>();
            table.AddRange(Product.GenerateProducts());
            table.AddColumn(t => t.Id);
            table.AddColumn(t => t.Name);
            table.AddColumn(t => t.Cost);
            table.AddColumn(t => t.LastPriceUpdate);


            string path = GetPath("BasicCsv");

            table.Create(path);
        }






        private string GetPath(string fileName)
        {
            DirectoryInfo? projectDir = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory)?.Parent?.Parent?.Parent;
            return $"{projectDir!.FullName}/Results/{fileName}";
        }

    }
}