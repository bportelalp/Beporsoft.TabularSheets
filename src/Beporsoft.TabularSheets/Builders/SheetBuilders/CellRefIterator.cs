namespace Beporsoft.TabularSheets.Builders.SheetBuilders
{
    /// <summary>
    /// Allow to iterate over a table, resetting and moving next row and col independently.
    /// </summary>
    public class CellRefIterator
    {
        private int _currentRow = 0;
        private int _currentCol = 0;
        private readonly bool zeroBasedIndex;


        /// <param name="zeroBasedIndex">If row and col starts by 0 or 1</param>
        public CellRefIterator(bool zeroBasedIndex = true)
        {
            this.zeroBasedIndex = zeroBasedIndex;
            Reset();
        }

        /// <summary>
        /// The current value of the iterator for the reference
        /// </summary>
        public string Current => CellRefBuilder.BuildRef(_currentRow, _currentCol, zeroBasedIndex);

        /// <summary>
        /// The current row value of the iterator for the reference
        /// </summary>
        public int CurrentRow => _currentRow;

        /// <summary>
        /// The current col value of the iterator for the reference as number
        /// </summary>
        public int CurrentCol => _currentCol;

        /// <summary>
        /// Increment the col iterator and return the new state
        /// </summary>
        public string MoveNextCol()
        {
            _currentCol++;
            return Current;
        }

        /// <summary>
        /// Increment the row iterator and return the new state
        /// </summary>
        public string MoveNextRow()
        {
            _currentRow++;
            return Current;
        }

        /// <summary>
        /// Increment the col iterator after return the previous state.
        /// </summary>
        public string MoveNextColAfter()
        {
            string current = Current;
            _currentCol++;
            return current;
        }

        /// <summary>
        /// Increment the row iterator after return the previous state.
        /// </summary>
        /// <returns></returns>
        public string MoveNextRowAfter()
        {
            string current = Current;
            _currentRow++;
            return current;
        }

        /// <summary>
        /// Restart the iterator of the columns
        /// </summary>
        public void ResetCol() => _currentCol = zeroBasedIndex? 0 : 1;

        /// <summary>
        /// Restart the iterator of the rows
        /// </summary>
        public void ResetRow() => _currentRow = zeroBasedIndex? 0 : 1;

        /// <summary>
        /// Restart the iterator
        /// </summary>
        public void Reset()
        {
            ResetCol();
            ResetRow();
        }
    }
}
