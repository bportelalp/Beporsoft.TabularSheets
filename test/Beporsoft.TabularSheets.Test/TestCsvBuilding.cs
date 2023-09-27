using Beporsoft.TabularSheets.Test.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Test
{
    [Category("CSV")]
    internal class TestCsvBuilding
    {
        private readonly bool _clearFolderOnEnd = false;
        private readonly TestFilesHandler _filesHandler = new TestFilesHandler("Csv");


        public void Start()
        {

        }

        [Test]
        public void ToCsv_Success()
        {
            string file = _filesHandler.BuildPath($"{nameof(ToCsv_Success)}.csv");
            var table = Product.GenerateTestSheet();
            table.ToCsv(file);
        }

        [OneTimeTearDown]
        public void OneTimeTearDown()
        {
            if (_clearFolderOnEnd)
                _filesHandler.ClearFiles();
        }
    }
}
