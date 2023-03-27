using Beporsoft.TabularSheets.Csv;
using Beporsoft.TabularSheets.Styling;
using System.Drawing;
using System.Globalization;

namespace Beporsoft.TabularSheets.Test.TestsTabularData
{
    public class TestTabularSheet
    {

        [Test]
        public void CheckFileName()
        {
            TabularSheet<Product> table = Generate();
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

        [Test]
        public void TestToDebug()
        {

            TabularSheet<Product> table = Generate();
            table.HeaderStyle.Fill.BackgroundColor = Color.Purple;
            table.DefaultStyle.Font.FontColor = Color.Red;
            table.HeaderStyle.Font.FontColor = Color.White;
            table.DefaultStyle.Fill.BackgroundColor = Color.AliceBlue;
            table.DefaultStyle.Border.SetAll(BorderStyle.BorderType.Thin);
            string path = GetPath("DebugTest.xlsx");
            table.Create(path);
        }


        private static TabularSheet<Product> Generate()
        {
            TabularSheet<Product> table = new();
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