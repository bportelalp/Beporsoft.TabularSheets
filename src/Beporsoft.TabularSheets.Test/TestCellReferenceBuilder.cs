using Beporsoft.TabularSheets.Builders.SheetBuilders;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Test
{
    /// <summary>
    /// Test class to create cell references
    /// </summary>
    internal class TestCellReferenceBuilder
    {

        [Test]
        [TestCaseSource(nameof(GetIndexColumnCases))]
        public void GetIndexColumn(string col, int expected)
        {
            int colNumber = CellRefBuilder.GetColIndex(col);
            Assert.That(colNumber, Is.EqualTo(expected));
        }

        [Test]
        [TestCaseSource(nameof(GetPositionCellFromReferenceCases))]
        public void GetPositionCellFromReference(string reference, int expectedRow, int expectedCol, bool zeroBasedIndex)
        {
            var position = CellRefBuilder.GetIndexes(reference, zeroBasedIndex);
            Assert.Multiple(() =>
            {
                Assert.That(position.Row, Is.EqualTo(expectedRow));
                Assert.That(position.Col, Is.EqualTo(expectedCol));
            });
        }


        [Test]
        [TestCaseSource(nameof(CreateCellReferenceCases))]
        public void CreateReference(string expected, int row, int col, bool zeroBasedIndex )
        {
            string reference = CellRefBuilder.BuildRef(row, col, zeroBasedIndex);
            Assert.That(reference, Is.EqualTo(expected));
        }

        private static IEnumerable<object> GetIndexColumnCases()
        {
            yield return new object[] { "A", 0 };
            yield return new object[] { "I", 8 };
            yield return new object[] { "DQ", 120 };
            yield return new object[] { "LP", 327 };
            yield return new object[] { "AAJ", 711 };
        }
        private static IEnumerable<object> GetPositionCellFromReferenceCases()
        {
            yield return new object[] { "A1", 0, 0, true };
            yield return new object[] { "A1", 1, 1, false };
            yield return new object[] { "DP300", 299, 119, true };
            yield return new object[] { "DP300", 300, 120, false };
        }
        private static IEnumerable<object> CreateCellReferenceCases()
        {
            yield return new object[] { "A1", 0, 0, true };
            yield return new object[] { "A1", 1, 1, false };


            yield return new object[] { "G6", 5, 6, true };
            yield return new object[] { "G6", 6, 7, false };

            yield return new object[] { "Z1", 0,25,  true };
            yield return new object[] { "Z1", 1, 26, false };

            yield return new object[] { "AA1", 0, 26, true };
            yield return new object[] { "AA1", 1, 27, false };

            yield return new object[] { "BF1", 0, 57, true };
            yield return new object[] { "BF1", 1, 58, false };

            yield return new object[] { "AAZ2", 1, 727, true };
            yield return new object[] { "AAZ2", 2, 728, false };
        }
    }
}
