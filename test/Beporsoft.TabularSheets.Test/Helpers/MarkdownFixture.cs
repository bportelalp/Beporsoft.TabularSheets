using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using Beporsoft.TabularSheets.Options;

namespace Beporsoft.TabularSheets.Test.Helpers
{
    internal class CsvFixture
    {
        private readonly CsvOptions _options;

        public CsvFixture(string path, CsvOptions options)
        {
            _options = options;
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);
            Load(fs, options);
        }

        public CsvFixture(Stream stream, CsvOptions options)
        {
            _options = options;
            Load(stream, options);

        }

        public string Header { get; private set; } = null!;
        public List<string> Lines { get; } = new();

        private void Load(Stream stream, CsvOptions options) {
            using var sr = new StreamReader(stream, options.Encoding);
            Header = sr.ReadLine()!;
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
            return Header.Split(_options.Separator)[colIndex];
        }

        public string GetCell(int row, int col)
        {
            string line = Lines[row];
            string cell = line.Split(_options.Separator)[col];
            return cell;
        }
    }
}
