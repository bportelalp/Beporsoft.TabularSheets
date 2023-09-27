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
        private readonly TestFilesHandler _filesHandler = new TestFilesHandler("SheetCreation");

        [Test, Category("StreamCreation")]
        [TestCaseSource(nameof(Data_FileHelpers_Verifypath))]
        public void FileHelpers_Verifypath_Ok(string? correctedPath, string providedPath, string[] targetExtensions)
        {
            string verify() => FileHelpers.VerifyPath(providedPath, targetExtensions);
            if (correctedPath is null)
            {
                // Must throw exception
                Assert.Throws<FileLoadException>(() => verify());
            }
            else
            {
                // Must give a result
                Assert.That(verify(), Is.EqualTo(correctedPath));
            }
        }


        [Test, Category("StreamCreation")]
        public void Create_CheckFileName()
        {
            TabularSheet<Product> table = Product.GenerateTestSheet();
            string pathOk = _filesHandler.BuildPath($"{nameof(Create_CheckFileName)}.xlsx");
            string pathNotOk = _filesHandler.BuildPath($"{nameof(Create_CheckFileName)}.xls");
            string pathWrongExtension = _filesHandler.BuildPath($"{nameof(Create_CheckFileName)}.csv");
            Assert.Multiple(() =>
            {
                Assert.That(() => table.Create(pathOk), Throws.Nothing);
                Assert.Catch<FileLoadException>(() => table.Create(pathNotOk));
                Assert.Catch<FileLoadException>(() => table.Create(pathWrongExtension));
            });
        }

        [OneTimeTearDown]
        public void OneTimeTearDown()
        {
            _filesHandler.ClearFiles();
        }

        private static IEnumerable<object?[]> Data_FileHelpers_Verifypath
        {
            get
            {
                yield return new object?[] { "File.xlsx", "File.xlsx", new string[] { ".xlsx", ".xls" } };
                yield return new object?[] { "File.xlsx", "File.xlsx", SpreadsheetFileExtension.AllowedExtensions };
                yield return new object?[] { "File.xls", "File.xls", new string[] { ".xlsx", ".xls" } };
                yield return new object?[] { "File.xlsx", "File", new string[] {".xlsx", ".xls" } };
                yield return new object?[] { "File.xls", "File", new string[] { ".xls", ".xlsx" } };
                yield return new object?[] { "File.csv", "File.csv", new string[] { ".xls", ".csv" } };
                yield return new object?[] { "File.csv", "File.csv", CsvFileExtension.AllowedExtensions };
                yield return new object?[] { "File.xlsx", "File", new string[] { ".xlsx", "csv" } };
                yield return new object?[] { null, "File.csv", new string[] { ".xlsx", ".xls" } };
            }
        }

    }
}
