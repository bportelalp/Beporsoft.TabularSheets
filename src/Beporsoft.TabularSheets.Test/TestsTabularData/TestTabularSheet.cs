using Beporsoft.TabularSheets.CellStyling;
using System.Drawing;
using System.Globalization;
using System.Reflection;

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
            Assert.That(() =>
            {
                TabularSheet<Product> table = Generate();
                table.HeaderStyle.Fill.BackgroundColor = Color.Purple;
                table.DefaultStyle.Font.Color = Color.Red;
                table.HeaderStyle.Font.Color = Color.White;
                table.DefaultStyle.Fill.BackgroundColor = Color.AliceBlue;
                table.DefaultStyle.Border.SetAll(BorderStyle.BorderType.Thin);
                string path = GetPath("DebugTest.xlsx");
                table.Create(path);
            }, Throws.Nothing);
        }

        [Test]
        public void TryHeaderStyles()
        {
            Assert.That(() =>
            {
                TabularSheet<Product> table = Generate();
                table.HeaderStyle.Fill.BackgroundColor = Color.Azure;
                //table.HeaderStyle.Font.Font = "Arial";
                table.HeaderStyle.Font.Size = 12;
                table.HeaderStyle.Border.SetAll(BorderStyle.BorderType.Medium);
                string path = GetPath($"Test{nameof(TryHeaderStyles)}.xlsx");
                table.Create(path);
            }, Throws.Nothing);
        }

        [Test]
        public void TryDefaultStyles()
        {
            Assert.That(() =>
            {
                TabularSheet<Product> table = Generate();
                table.DefaultStyle.Fill.BackgroundColor = Color.Azure;
                table.DefaultStyle.Font.Font = "Arial";
                table.DefaultStyle.Font.Size = 8;
                table.DefaultStyle.Border.SetAll(BorderStyle.BorderType.Thin);
                string path = GetPath($"Test{nameof(TryDefaultStyles)}.xlsx");
                table.Create(path);
            }, Throws.Nothing);
        }

        [Test]
        public void TryOverrideStyle()
        {
            Assert.That(() =>
            {
                TabularSheet<Product> table = Generate();
                table.DefaultStyle.Fill.BackgroundColor = Color.Azure;
                table.DefaultStyle.Font.Font = "Arial";
                table.DefaultStyle.Font.Size = 8;
                table.DefaultStyle.Border.SetAll(BorderStyle.BorderType.Thin);

                table.HeaderStyle.Fill.BackgroundColor = Color.DarkBlue;
                table.HeaderStyle.Font.Color = Color.White;
                table.HeaderStyle.Border.SetAll(BorderStyle.BorderType.Medium);
                table.HeaderStyle.Border.Color = Color.Yellow;
                string path = GetPath($"Test{nameof(TryOverrideStyle)}.xlsx");
                table.Create(path);
            }, Throws.Nothing);
        }

        [Test]
        public void TryColumnStyle()
        {
            Assert.That(() =>
            {
                Style style = new Style();
                style.Fill.BackgroundColor = Color.Red;
                TabularSheet<Product> table = Generate();
                table.AddColumn(t => t.Name).SetTitle("Name duplicated").SetStyle(style);
                string path = GetPath($"Test{nameof(TryColumnStyle)}.xlsx");
                table.Create(path);
            }, Throws.Nothing);
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