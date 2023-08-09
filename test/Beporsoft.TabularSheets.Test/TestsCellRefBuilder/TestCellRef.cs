using Beporsoft.TabularSheets.Builders.SheetBuilders;

namespace Beporsoft.TabularSheets.Test.TestsCellRefBuilder
{
    /// <summary>
    /// Test class to create cell references
    /// </summary>
    internal class TestCellRef
    {

        [Test]
        [TestCaseSource(nameof(SetOfCellRef))]
        public void GetIndexes(string reference, int expectedRow, int expectedCol, bool zeroBasedIndex)
        {
            var (Row, Col) = CellRefBuilder.GetIndexes(reference, zeroBasedIndex);
            Assert.Multiple(() =>
            {
                Assert.That(Row, Is.EqualTo(expectedRow));
                Assert.That(Col, Is.EqualTo(expectedCol));
            });
        }


        [Test]
        [TestCaseSource(nameof(SetOfCellRef))]
        public void CreateReference(string expected, int row, int col, bool zeroBasedIndex)
        {
            string reference = CellRefBuilder.BuildRef(row, col, zeroBasedIndex);
            Assert.That(reference, Is.EqualTo(expected));
        }

        [Test]
        [TestCaseSource(nameof(SetOfCellRef))]
        public void TestCellIterator(string expected, int row, int col, bool zeroBasedIndex)
        {
            int iterationsRow = zeroBasedIndex ? row : row - 1;
            int iterationsCol = zeroBasedIndex ? col : col - 1;
            var iterator = new CellRefIterator(zeroBasedIndex);
            // Iterate over row and then over col
            for (int i = 0; i < iterationsRow; i++)
            {
                iterator.MoveNextRow();
            }
            for (int i = 0; i < iterationsCol; i++)
            {
                iterator.MoveNextCol();

            }
            Assert.That(iterator.Current, Is.EqualTo(expected));
            // Check reset
            iterator.Reset();
            Assert.That(iterator.Current, Is.EqualTo("A1"));
        }

        /// <summary>
        /// Test cases [ref, expectedRow, expectedCol, zeroBasedIndex]
        /// </summary>
        private static IEnumerable<object> SetOfCellRef()
        {
            yield return new object[] { "A1", 0, 0, true };
            yield return new object[] { "A1", 1, 1, false };


            yield return new object[] { "G6", 5, 6, true };
            yield return new object[] { "G6", 6, 7, false };

            yield return new object[] { "DP300", 299, 119, true };
            yield return new object[] { "DP300", 300, 120, false };

            yield return new object[] { "Z1", 0, 25, true };
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
