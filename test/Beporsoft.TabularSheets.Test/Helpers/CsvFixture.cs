using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using Beporsoft.TabularSheets.Options;

namespace Beporsoft.TabularSheets.Test.Helpers
{
    internal class MarkdownFixture
    {

        public MarkdownFixture(string path, MarkdownTableOptions? options)
        {
            Options = options ?? MarkdownTableOptions.Default;
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);
            Load(fs);
        }

        public MarkdownFixture(Stream stream, MarkdownTableOptions? options)
        {
            Options = options ?? MarkdownTableOptions.Default;
            Load(stream);
        }

        public MarkdownFixture(IEnumerable<string> lines, MarkdownTableOptions? options)
        {
            Options = options ?? MarkdownTableOptions.Default;
            Header = lines.First();
            HeaderSeparator = lines.Skip(1).First();
            Lines = lines.Skip(2).ToList();
        }

        public string Header { get; private set; } = null!;
        public string HeaderSeparator { get; private set; } = null!;
        public List<string> Lines { get; } = new();
        public MarkdownTableOptions Options { get; }

        private void Load(Stream stream) {
            using var sr = new StreamReader(stream);
            Header = sr.ReadLine()!;
            HeaderSeparator = sr.ReadLine()!;
            string? line = null;
            do
            {
                line = sr.ReadLine();
                if (line is not null)
                    Lines.Add(line);
            } while (line is not null);
        }

        public string GetHeaderColumn(int colIndex)
        {
            var split = Header.Trim('|').Split(MarkdownTableOptions.Separator);
            return split[colIndex].Trim();
        }

        public int GetHeaderLenght(int colIndex)
        {
            var split = Header.Trim('|').Split(MarkdownTableOptions.Separator);
            return split[colIndex].Length;
        }

        public string GetHeaderSeparatorColumn(int colIndex)
        {
            var split = HeaderSeparator.Trim('|').Split(MarkdownTableOptions.Separator);
            return split[colIndex].Trim();
        }

        public int GetHeaderSeparatorLenght(int colIndex)
        {
            var split = HeaderSeparator.Trim('|').Split(MarkdownTableOptions.Separator);
            return split[colIndex].Length;
        }


        public string GetCell(int row, int col)
        {
            string line = Lines[row];
            string cell = line.Trim('|').Split(MarkdownTableOptions.Separator)[col].Trim();
            return cell;
        }

        public int GetCellLength(int row, int col)
        {
            string line = Lines[row];
            string cell = line.Trim('|').Split(MarkdownTableOptions.Separator)[col];
            return cell.Length;
        }
    }
}
