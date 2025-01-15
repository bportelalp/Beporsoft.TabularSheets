using Beporsoft.TabularSheets.Options;
using System.Collections.Generic;

namespace Beporsoft.TabularSheets.Builders
{
    internal sealed class MarkdownTableBuilder<T>
    {
        private const string _separator = "|";
        private const string _whiteSpace = " ";
        private const string _headerLineContent = "--";
        public MarkdownTableBuilder(TabularData<T> tabularData, MarkdownParsingOptions options)
        {
            TabularData = tabularData;
            Options = options;
        }

        public TabularData<T> TabularData { get; }
        public MarkdownParsingOptions Options { get; }

        public List<string> Create()
        {
            List<string> result = [..CreateHeader()];
            foreach (var row in TabularData)
            {
                string mdRow = CreateLine(row);
                result.Add(mdRow);
            }
            return result;
        }

        private List<string> CreateHeader()
        {
            string header = _separator;
            string headerBodySeparator = _separator;
            foreach (var column in TabularData.Columns)
            {
                header += _whiteSpace + column.Title + _whiteSpace;
                header += _separator;
                headerBodySeparator += _whiteSpace + _headerLineContent + _whiteSpace + _separator;
            }
            return [header, headerBodySeparator];
        }

        private string CreateLine(T row)
        {
            string line = _separator;
            foreach (var column in TabularData.Columns)
            {
                line += _whiteSpace + column.Apply(row) + _whiteSpace;
                line += _separator;
            }
            return line;
        }

    }
}
