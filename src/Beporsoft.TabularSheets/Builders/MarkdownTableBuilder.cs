using Beporsoft.TabularSheets.Options;
using System.Collections.Generic;
using System.IO;

namespace Beporsoft.TabularSheets.Builders
{
    internal sealed class MarkdownTableBuilder<T>
    {
        private const string _separator = "|";
        private const int _minColumnPadding = 2;
        private readonly Dictionary<int, int> _colWidths = [];

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

            foreach (string row in CreateHeader())
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
            using FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write);
            WriteStream(fs);
        }

        public void WriteStream(Stream stream)
        {
            using StreamWriter sw = new StreamWriter(stream);
            foreach (string mdRow in Create())
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
                string value = string.Empty;
                if (!Options.SupressHeaderTitles)
                    value = PrintValue(column.Title, column.Index);
                header += $" {value} {_separator}";
                headerBodySeparator += $" {PrintValue("--", column.Index, '-')} {_separator}";
            }
            return [header, headerBodySeparator];
        }

        private string CreateLine(T row)
        {
            string line = _separator;
            foreach (var column in TabularData.Columns)
            {
                string value = column.Apply(row)?.ToString() ?? string.Empty;
                line += $" {PrintValue(value, column.Index)} ";
                line += _separator;
            }
            return line;
        }

        private void CalculateColumnLengths()
        {
            _colWidths.Clear();
            foreach (var column in TabularData.Columns)
            {
                _colWidths[column.Index] = column.Title.Length < _minColumnPadding ? _minColumnPadding : column.Title.Length;
                foreach (var row in TabularData.Items)
                {
                    object value = column.Apply(row);
                    int stringLength = value?.ToString()?.Length ?? 0;
                    if (stringLength > _colWidths[column.Index])
                        _colWidths[column.Index] = stringLength;
                }
            }
        }

        private string PrintValue(string value, int columnIndex, char padChar = ' ')
        {
            if (Options.CompactMode)
                return value;

            if (!_colWidths.TryGetValue(columnIndex, out int length))
                return value;

            return value.PadRight(length, padChar);
        }

    }
}
