using Beporsoft.TabularSheet.Csv;

namespace Beporsoft.TabularSheet.Test
{
    public class TestTabularCsv
    {

        [Test]
        public void Creation()
        {
            string path = GetPath("BasicCsv.csv");
            TabularCsv<Product> table = new TabularCsv<Product>();
            table.AddRange(Product.GenerateProducts());
            table.SetColumn(t => t.Id);
            table.SetColumn(t => t.Name);
            table.SetColumn(t => t.Cost);
            table.SetColumn(t => t.LastPriceUpdate);
            table.Create(path);
        }

        [Test]
        public void CreationPathWrong()
        {
            TabularCsv<Product> table = new TabularCsv<Product>();
            table.AddRange(Product.GenerateProducts());
            table.SetColumn(t => t.Id);
            table.SetColumn(t => t.Name);
            table.SetColumn(t => t.Cost);
            table.SetColumn(t => t.LastPriceUpdate);


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