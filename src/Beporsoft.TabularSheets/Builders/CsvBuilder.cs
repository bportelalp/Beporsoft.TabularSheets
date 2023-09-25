using Beporsoft.TabularSheets.Builders.Interfaces;
using Beporsoft.TabularSheets.Csv;
using Beporsoft.TabularSheets.Tools;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Builders
{
    internal sealed class CsvBuilder<T>
    {
        public TabularData<T> TabularData { get; }
        public CsvOptions Options { get; }

        public CsvBuilder(TabularData<T> tabularData, CsvOptions options)
        {
            TabularData = tabularData;
            Options = options;
        }

        public void Create(string path)
        {
            string pathCorrected = FileHelpers.VerifyPath(path, CsvFileExtension.Csv);
            using var fs = new FileStream(pathCorrected, FileMode.Create);
            using MemoryStream ms = Create();
            ms.Seek(0, SeekOrigin.Begin);
            ms.CopyTo(fs);
        }

        public MemoryStream Create()
        {
            MemoryStream stream = new MemoryStream();
            StreamWriter writer = new StreamWriter(stream, Options.Encoding);
            foreach (var line in CreateLines())
            {
                writer.WriteLine(line);
            }
            writer.Flush();
            stream.Seek(0, SeekOrigin.Begin);
            return stream;
        }

        #region Create Lines
        private IEnumerable<string> CreateLines()
        {
            string header = CreateHeader();
            yield return header;
            foreach (var item in TabularData)
            {
                yield return CreateLine(item);
            }
        }

        private string CreateHeader()
        {
            string header = string.Empty;
            foreach (var column in TabularData.Columns)
            {
                header += column.Title;
                if (column.Index != TabularData.ColumnCount - 1)
                    header += Options.Separator;
            }
            return header;
        }
        private string CreateLine(T row)
        {
            string line = string.Empty;
            foreach (var column in TabularData.Columns)
            {
                line += column.Apply(row);
                if (column.Index != TabularData.ColumnCount - 1)
                    line += Options.Separator;
            }
            return line;
        }
        #endregion
    }
}
