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
        public MarkdownTableBuilder(TabularData<T> tabularData, MarkdownTableOptions? options)
        {
            TabularData = tabularData;
            Options = options?? MarkdownTableOptions.Default;
        }

        public TabularData<T> TabularData { get; }
        public MarkdownTableOptions Options { get; }

        public IEnumerable<string> Create()
        {
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
                    header += column.Title;
                header += _whiteSpace;
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
