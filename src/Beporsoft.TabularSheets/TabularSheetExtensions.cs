using Beporsoft.TabularSheets.Builders;
using Beporsoft.TabularSheets.Csv;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets
{
    public static class TabularSheetExtensions
    {

        public static MemoryStream ToCsv<T>(this TabularData<T> tabularData) => tabularData.ToCsv(new CsvOptions());

        public static MemoryStream ToCsv<T>(this TabularData<T> tabularData, CsvOptions options)
        {
            CsvBuilder<T> builder = new CsvBuilder<T>(tabularData, options);
            MemoryStream ms = builder.Create();
            return ms;
        }

        public static void ToCsv<T>(this TabularData<T> tabularData, string path) => tabularData.ToCsv(path, new CsvOptions());

        public static void ToCsv<T>(this TabularData<T> tabularData, string path, CsvOptions options)
        {
            CsvBuilder<T> builder = new CsvBuilder<T>(tabularData, options);
            builder.Create(path);
        }
    }
}
