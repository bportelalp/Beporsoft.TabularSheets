using Beporsoft.TabularSheets.Builders.SheetBuilders;
using Beporsoft.TabularSheets.Test.Helpers;
using Beporsoft.TabularSheets.Test.TestModels;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Test
{
    [Category("SheetStructure"), Category("AutoFilter")]
    public class TestAutoFilter
    {
        private readonly bool _clearFolderOnEnd = false;
        private readonly int _amountRows = 1000;
        private readonly TestFilesHandler _filesHandler = new(nameof(TestAutoFilter));
        private readonly Stopwatch _stopwatch = new Stopwatch();

        [SetUp]
        public void Setup()
        {
            _stopwatch.Reset();
        }

        [OneTimeTearDown]
        public void OneTimeTearDown()
        {
            if (_clearFolderOnEnd)
                _filesHandler.ClearFiles();

        }

        [Test]
        public void Create_NotActivatedAutoFilter_NotIncludedOnSheet()
        {
            string path = _filesHandler.BuildPath($"Test{nameof(Create_NotActivatedAutoFilter_NotIncludedOnSheet)}.xlsx");
            TabularSheet<Product> table = Product.GenerateProductSheet(_amountRows);
            table.Options.UseAutofilter = false;
            table.Create(path);
            SheetFixture sheet = new SheetFixture(path);

            AutoFilter? autoFilter = sheet.AutoFilter;
            Assert.That(autoFilter, Is.Null);
        }

        [Test]
        public void Create_ActivatedAutoFilter_IncludedAndRangeOk()
        {
            string path = _filesHandler.BuildPath($"Test{nameof(Create_ActivatedAutoFilter_IncludedAndRangeOk)}.xlsx");

            TabularSheet<Product> table = Product.GenerateProductSheet();
            table.Options.UseAutofilter = true;
            table.Create(path);

            SheetFixture sheet = new SheetFixture(path);

            AutoFilter? autoFilter = sheet.AutoFilter;
            SheetDimension dimension = sheet.Dimensions;
            Assert.That(autoFilter, Is.Not.Null);

            string rangeDimensions = CellRefBuilder.BuildRefRange(CellRefBuilder.BuildRef(0, 0), CellRefBuilder.BuildRef(table.Count, table.ColumnCount, false));
            Assert.Multiple(() =>
            {
                Assert.That(autoFilter.Reference, Is.Not.Null);
                Assert.That(autoFilter.Reference!.Value, Is.EqualTo(rangeDimensions));

                Assert.That(dimension.Reference, Is.Not.Null);
                Assert.That(dimension.Reference!.Value, Is.EqualTo(rangeDimensions));

                Assert.That(autoFilter.Reference, Is.EqualTo(dimension.Reference));
            });
        }
    }
}
