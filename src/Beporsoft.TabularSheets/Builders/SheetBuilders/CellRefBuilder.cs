using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Beporsoft.TabularSheets.Builders.SheetBuilders
{
    /// <summary>
    /// A class to build cell reference values from row and col indexes
    /// </summary>
    public static class CellRefBuilder
    {
        internal const string AllowedChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        internal static readonly Regex RegexReference = new(@"^\D{1,}\d{1,}$");
        internal static readonly Regex RegexReferenceRow = new(@"\d{1,}$");
        internal static readonly Regex RegexReferenceColumn = new(@"^\D{1,}");

        /// <summary>
        /// Build the reference of a cell based on row and col index
        /// </summary>
        /// <param name="row">Number of the row</param>
        /// <param name="col">Number of the column</param>
        /// <param name="zeroBasedIndex">Whether start index on <paramref name="row"/> and <paramref name="col"/> is 0 or 1</param>
        public static string BuildRef(int row, int col, bool zeroBasedIndex = true)
        {
            string colRef = BuildColRef(col, zeroBasedIndex);
            string rowRef = BuildRowRef(row, zeroBasedIndex);
            string result = colRef + rowRef;
            return result;
        }

        /// <summary>
        /// Build the col part reference of a cell based on it col index
        /// </summary>
        /// <param name="col"></param>
        /// <param name="zeroBasedIndex">Whether start index on row and <paramref name="col"/> is 0 or 1</param>
        public static string BuildColRef(int col, bool zeroBasedIndex = true)
        {
            List<int> indexes = new();
            int operand = zeroBasedIndex ? col : col - 1;
            while (operand >= 26)
            {
                int rest = operand % 26;
                indexes.Add(rest);
                operand = (operand - rest - 1) / 26;
            }
            indexes.Add(operand);

            indexes.Reverse();

            string colReference = string.Empty;
            foreach (var index in indexes)
                colReference += AllowedChars[index];
            return colReference;
        }

        /// <summary>
        /// Build the row part reference of a cell based on it col index
        /// </summary>
        /// <param name="row"></param>
        /// <param name="zeroBasedIndex">Whether start index on <paramref name="row"/> and col is 0 or 1</param>
        public static string BuildRowRef(int row, bool zeroBasedIndex = true)
        {
            int rowReference = zeroBasedIndex ? row + 1 : row;
            string result = rowReference.ToString();
            return result;
        }

        /// <summary>
        /// Get the col index from reference
        /// </summary>
        /// <param name="cellReference">The alphanumeric reference of the cell</param>
        /// <param name="zeroBasedIndex">Whether start index on row and col is 0 or 1</param>
        public static int GetColIndex(string cellReference, bool zeroBasedIndex = true)
        {
            string colRef = GetColPart(cellReference);
            int sum = 0;
            for (int i = 0; i < colRef.Length; i++)
            {
                sum *= 26;
                sum += colRef[i] - 'A' + 1;
            }
            return zeroBasedIndex ? sum - 1 : sum;
        }

        /// <summary>
        /// Get the row index from reference
        /// </summary>
        /// <param name="cellReference">The alphanumeric reference of the cell</param>
        /// <param name="zeroBasedIndex">Whether start index on row and col is 0 or 1</param>
        public static int GetRowIndex(string cellReference, bool zeroBasedIndex = true)
        {
            string rowRef = GetRowPart(cellReference);
            int row = Convert.ToInt32(rowRef);
            return zeroBasedIndex ? row - 1 : row;
        }

        /// <summary>
        /// Get the indexes of a cell reference
        /// </summary>
        /// <param name="cellReference">The alphanumeric reference of the cell</param>
        /// <param name="zeroBasedIndex">Whether start index on row and col is 0 or 1</param>
        public static (int Row, int Col) GetIndexes(string cellReference, bool zeroBasedIndex = true)
        {
            int row = GetRowIndex(cellReference, zeroBasedIndex);
            int col = GetColIndex(cellReference, zeroBasedIndex);
            return (row, col);
        }

        private static void CheckFormat(string reference)
        {
            Match match = RegexReference.Match(reference);
            if (!match.Success)
                throw new ArgumentException($"The cell reference isn't in the correct format. Value={reference}");
        }

        private static string GetColPart(string reference)
        {
            CheckFormat(reference);
            Match match = RegexReferenceColumn.Match(reference);
            return match.Value;
        }

        private static string GetRowPart(string reference)
        {
            CheckFormat(reference);
            Match match = RegexReferenceRow.Match(reference);
            return match.Value;
        }
    }
}
