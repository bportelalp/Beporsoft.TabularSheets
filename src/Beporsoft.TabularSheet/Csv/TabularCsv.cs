using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheet.Csv
{
    public class TabularCsv<T> : TabularData<T>, ITabularDataStorage
    {
        private const string _defaultExtension = ".csv";
        public CsvDelimiter Delimiter { get; set; }

        public void Create(string path) => Create(path, Encoding.GetEncoding("latin1"));
        public void Create(string path, Encoding encoding)
        {
            VerifyPath(path);
            using var fs = new FileStream(path, FileMode.Create, FileAccess.Write);
            using MemoryStream ms = Create(encoding);
            ms.Seek(0, SeekOrigin.Begin);
            ms.CopyTo(fs);
        }

        public MemoryStream Create() => Create(Encoding.GetEncoding("latin1"));
        public MemoryStream Create(Encoding encoding)
        {
            List<string> lines = CreateLines().ToList();
            MemoryStream stream = new MemoryStream();
            StreamWriter writer = new StreamWriter(stream, encoding);
            foreach (var line in lines)
            {
                writer.WriteLine(line);
            }
            writer.Flush();
            return stream;
        }

        private static string VerifyPath(string path)
        {
            bool hasExtension = Path.HasExtension(path);
            if (hasExtension)
            {
                string extension = Path.GetExtension(path).Trim();
                if (extension != _defaultExtension)
                    throw new FileLoadException($"The path provided does not aim a {_defaultExtension} file extension", path);
                else
                    return path;
            }
            else
            {
                return $"{path}{_defaultExtension}";
            }
        }

        #region Create Lines
        private IEnumerable<string> CreateLines()
        {
            string header = CreateHeader();
            yield return header;
            foreach (var item in this)
            {
                yield return CreateLine(item);
            }
        }

        private string CreateHeader()
        {
            string header = string.Empty;
            foreach (var column in Columns)
            {
                header += column.Title;
                if (Columns.Last() != column)
                    header += Delimiter.GetChar();
            }
            return header;
        }
        private string CreateLine(T row)
        {
            string line = string.Empty;
            foreach (var column in Columns)
            {
                line += column.Apply(row);
                if (Columns.Last() != column)
                    line += Delimiter.GetChar();
            }
            return line;
        }
        #endregion
    }
}
