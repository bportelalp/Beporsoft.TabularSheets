using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Builders.SheetBuilders
{
    /// <summary>
    /// A class to build cell reference values from row and col indexes
    /// </summary>
    internal static class CellReferenceBuilder
    {
        internal const string AllowedChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        internal static readonly Regex RegexReference = new(@"^\D{1,}\d{1,}$");
        internal static readonly Regex RegexReferenceRow = new(@"\d{1,}$");
        internal static readonly Regex RegexReferenceColumn = new(@"^\D{1,}");

        internal static string Build(int row, int col, bool zeroBasedIndex = true)
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

            int rowReference = zeroBasedIndex ? row + 1 : row;
            string result = colReference + rowReference.ToString();
            return result;
        }

        public static int GetColumnIndex(string colReference, bool zeroBasedIndex = true)
        {
            colReference = colReference.Trim().ToUpper();

            int sum = 0;

            for (int i = 0; i < colReference.Length; i++)
            {
                sum *= 26;
                sum += colReference[i] - 'A' + 1;
            }
            return HandleZeroIndex(sum, zeroBasedIndex);
        }

        internal static (int Row, int Col) GetIndexes(string reference, bool zeroBasedIndex = true)
        {
            int row = GetRowPart(reference);
            string colStr = GetColumnPart(reference);
            int col = GetColumnIndex(colStr, false);
            return (HandleZeroIndex(row, zeroBasedIndex), HandleZeroIndex(col, zeroBasedIndex));
        }



        internal static void CheckFormat(string reference)
        {
            Match match = RegexReference.Match(reference);
            if (!match.Success)
                throw new ArgumentException($"The cell reference isn't in the correct format. Value={reference}");
        }

        private static string GetColumnPart(string reference)
        {
            CheckFormat(reference);
            Match match = RegexReferenceColumn.Match(reference);
            return match.Value;
        }

        private static int GetRowPart(string reference)
        {
            CheckFormat(reference);
            Match match = RegexReferenceRow.Match(reference);
            return Convert.ToInt32(match.Value);
        }

        private static int HandleZeroIndex(int value, bool zeroBasedIndex)
        {
            if (zeroBasedIndex)
                return value - 1;
            else
                return value;
        }
    }
}
