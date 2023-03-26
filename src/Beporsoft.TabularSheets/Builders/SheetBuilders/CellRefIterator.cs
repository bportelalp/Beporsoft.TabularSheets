using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Builders.SheetBuilders
{
    /// <summary>
    /// Allow to iterate over a table, resetting and moving next row and col independently.
    /// </summary>
    public class CellRefIterator
    {
        private int _currentRow = 0;
        private int _currentCol = 0;
        public string Current => CellRefBuilder.BuildRef(_currentRow, _currentCol);

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


        public void ResetCol() => _currentCol = 0;
        public void ResetRow() => _currentRow = 0;
        public void Reset()
        {
            _currentRow = 0;
            _currentCol = 0;
        }
    }
}
