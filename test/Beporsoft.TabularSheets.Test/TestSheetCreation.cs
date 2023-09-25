using Beporsoft.TabularSheets.Test.Helpers;
using Beporsoft.TabularSheets.Tools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Test
{
    
    public class TestSheetCreation
    {



        [Test, Category("StreamCreation")]
        [TestCaseSource(nameof(Data_FileHelpers_Verifypath_Ok))]
        public void FileHelpers_Verifypath_Ok(string correctedPath, string providedPath, string[] targetExtensions)
        {
            string result = FileHelpers.VerifyPath(providedPath, targetExtensions);
            Assert.That(result, Is.EqualTo(correctedPath));
        }


        [Test, Category("StreamCreation")]
        public void Create_CheckFileName()
        {
            TabularSheet<Product> table = Product.GenerateTestSheet();
            string pathOk = TestDirectory.GetPath($"{nameof(Create_CheckFileName)}.xlsx");
            string pathNotOk = TestDirectory.GetPath($"{nameof(Create_CheckFileName)}.xls");
            string pathWrongExtension = TestDirectory.GetPath($"{nameof(Create_CheckFileName)}.csv");
            Assert.Multiple(() =>
            {
                Assert.That(() => table.Create(pathOk), Throws.Nothing);
                Assert.Catch<FileLoadException>(() => table.Create(pathNotOk));
                Assert.Catch<FileLoadException>(() => table.Create(pathWrongExtension));
            });
        }

        private static IEnumerable<object[]> Data_FileHelpers_Verifypath_Ok
        {
            get
            {
                yield return new object[] { "File.xlsx", "File.xlsx", new string[] { ".xlsx", ".xls" } };
                yield return new object[] { "File.xls", "File.xls", new string[] { ".xlsx", ".xls" } };
                yield return new object[] { "File.xlsx", "File", new string[] {".xlsx", ".xls" } };
                yield return new object[] { "File.xls", "File", new string[] { ".xls", ".xlsx" } };
            }
        }
    }
}
