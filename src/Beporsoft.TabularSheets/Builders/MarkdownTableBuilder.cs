using Beporsoft.TabularSheets.Options;
using System.Collections.Generic;
using System.IO;

namespace Beporsoft.TabularSheets.Builders
{
    internal sealed class MarkdownTableBuilder<T>
    {
        private const string _separator = "|";
        private const string _whiteSpace = " ";
        private const string _headerLineContent = "--";
        private const int _minColumnPadding = 2;
        private readonly Dictionary<int, int> _lengths = [];

        public MarkdownTableBuilder(TabularData<T> tabularData, MarkdownTableOptions? options)
        {
            TabularData = tabularData;
            Options = options ?? MarkdownTableOptions.Default;
        }

        public TabularData<T> TabularData { get; }
        public MarkdownTableOptions Options { get; }

        public IEnumerable<string> Create()
        {
            if (!Options.CompactMode)
                CalculateColumnLengths();

            foreach (var row in CreateHeader())
            {
                yield return row;
            }
            foreach (var row in TabularData)
            {
                string mdRow = CreateLine(row);
                yield return mdRow;
            }
        }

        public void Create(string path)
        {
            using var fs = new FileStream(path, FileMode.Create, FileAccess.Write);
            WriteStream(fs);
        }

        public void WriteStream(Stream stream)
        {
            using var sw = new StreamWriter(stream);
            foreach (var mdRow in Create())
            {
                sw.WriteLine(mdRow);
            }
        }

        private IEnumerable<string> CreateHeader()
        {
            string header = _separator;
            string headerBodySeparator = _separator;
            foreach (var column in TabularData.Columns)
            {
                header += _whiteSpace;
                if (!Options.SupressHeaderTitles)
                    header += PrintValue(column.Title, column.Index);
                header += _whiteSpace;
                header += _separator;
                headerBodySeparator += _whiteSpace + PrintValue(_headerLineContent, column.Index, '-') + _whiteSpace + _separator;
            }
            return [header, headerBodySeparator];
        }

        private string CreateLine(T row)
        {
            string line = _separator;
            foreach (var column in TabularData.Columns)
            {
                string value = column.Apply(row)?.ToString() ?? string.Empty;
                line += _whiteSpace + PrintValue(value, column.Index) + _whiteSpace;
                line += _separator;
            }
            return line;
        }

        private void CalculateColumnLengths()
        {
            _lengths.Clear();
            foreach (var column in TabularData.Columns)
            {
                _lengths[column.Index] = column.Title.Length < _minColumnPadding ? _minColumnPadding : column.Title.Length;
                foreach (var row in TabularData.Items)
                {
                    object value = column.Apply(row);
                    int stringLength = value?.ToString()?.Length ?? 0;
                    if (stringLength > _lengths[column.Index])
                        _lengths[column.Index] = stringLength;
                }
            }
        }

        private string PrintValue(string value, int columnIndex, char padChar = ' ')
        {
            if (Options.CompactMode)
                return value;

            if (!_lengths.TryGetValue(columnIndex, out var length))
                return value;

            return value.PadRight(length, padChar);
        }

    }
}
