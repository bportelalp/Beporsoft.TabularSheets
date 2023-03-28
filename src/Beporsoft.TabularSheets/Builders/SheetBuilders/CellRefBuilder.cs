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
        /// <param name="zeroBasedIndex">Whether start index on <paramref name="row"/> and <paramref name="col"/> is 0 or 1</param>
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
        /// <param name="zeroBasedIndex">Whether start index on <paramref name="row"/> and <paramref name="col"/> is 0 or 1</param>
        public static string BuildRowRef(int row, bool zeroBasedIndex = true)
        {
            int rowReference = zeroBasedIndex ? row + 1 : row;
            string result = rowReference.ToString();
            return result;
        }

        /// <summary>
        /// Get the col index from reference
        /// </summary>
        /// <param name="zeroBasedIndex">Whether start index on <paramref name="row"/> and <paramref name="col"/> is 0 or 1</param>
        public static int GetColIndex(string reference, bool zeroBasedIndex = true)
        {
            string colRef = GetColPart(reference);
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
        /// <param name="zeroBasedIndex">Whether start index on <paramref name="row"/> and <paramref name="col"/> is 0 or 1</param>
        public static int GetRowIndex(string reference, bool zeroBasedIndex = true)
        {
            string rowRef = GetRowPart(reference);
            int row = Convert.ToInt32(rowRef);
            return zeroBasedIndex ? row - 1 : row;
        }

        /// <summary>
        /// Get the indexes of a cell reference
        /// </summary>
        /// <param name="zeroBasedIndex">Whether start index on <paramref name="row"/> and <paramref name="col"/> is 0 or 1</param>
        public static (int Row, int Col) GetIndexes(string cellRef, bool zeroBasedIndex = true)
        {
            int row = GetRowIndex(cellRef, zeroBasedIndex);
            int col = GetColIndex(cellRef, zeroBasedIndex);
            return (row, col);
        }

        private static void CheckFormat(string reference)
        {
            Match match = RegexReference.Match(reference);
            if (!match.Success)
                throw new ArgumentException($"The cell reference isn't in the correct format. Value={reference}");
        }

        public static string GetColPart(string reference)
        {
            CheckFormat(reference);
            Match match = RegexReferenceColumn.Match(reference);
            return match.Value;
        }

        public static string GetRowPart(string reference)
        {
            CheckFormat(reference);
            Match match = RegexReferenceRow.Match(reference);
            return match.Value;
        }
    }
}
