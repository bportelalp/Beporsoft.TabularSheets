using Beporsoft.TabularSheets.Builders;
using Beporsoft.TabularSheets.Options;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets
{
    /// <summary>
    /// Extension methods for <see cref="TabularData{T}"/>.
    /// </summary>
    public static class TabularDataExtensions
    {
        #region ToCsv
        /// <summary>
        /// Creates a CSV file as stream.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="tabularData"></param>
        /// <returns>A stream which contains the csv file</returns>
        public static MemoryStream ToCsv<T>(this TabularData<T> tabularData) => tabularData.ToCsv(new CsvOptions());

        /// <summary>
        /// <inheritdoc cref="ToCsv{T}(TabularData{T})"/>
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="tabularData"></param>
        /// <param name="options">Configurations for the generation</param>
        /// <returns>A stream which contains the csv file</returns>
        public static MemoryStream ToCsv<T>(this TabularData<T> tabularData, CsvOptions options)
        {
            CsvBuilder<T> builder = new CsvBuilder<T>(tabularData, options);
            MemoryStream ms = builder.Create();
            return ms;
        }

        /// <summary>
        /// Creates a CSV file writting it on the given stream.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="tabularData"></param>
        /// <param name="stream"></param>
        public static void ToCsv<T>(this TabularData<T> tabularData, Stream stream) => tabularData.ToCsv(stream, new CsvOptions());

        /// <summary>
        /// <inheritdoc cref="ToCsv{T}(TabularData{T}, Stream)"/>
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="tabularData"></param>
        /// <param name="stream"></param>
        /// <param name="options">Configurations for the generation</param>
        public static void ToCsv<T>(this TabularData<T> tabularData, Stream stream, CsvOptions options)
        {
            CsvBuilder<T> builder = new CsvBuilder<T>(tabularData, options);
            builder.Create(stream);
        }

        /// <summary>
        /// Creates a CSV file on the given path.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="tabularData"></param>
        /// <param name="path"></param>
        public static void ToCsv<T>(this TabularData<T> tabularData, string path) => tabularData.ToCsv(path, new CsvOptions());

        /// <summary>
        /// <inheritdoc cref="ToCsv{T}(TabularData{T}, string)"/>
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="tabularData"></param>
        /// <param name="path"></param>
        /// <param name="options">Configurations for the generation</param>
        public static void ToCsv<T>(this TabularData<T> tabularData, string path, CsvOptions options)
        {
            CsvBuilder<T> builder = new CsvBuilder<T>(tabularData, options);
            builder.Create(path);
        }
        #endregion

        #region ToMarkdown
        /// <summary>
        /// Creates a Markdown formatted table on the given path.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="tabularData"></param>
        /// <param name="path"></param>
        public static void ToMarkdownTable<T>(this TabularData<T> tabularData, string path, MarkdownTableOptions? mdOptions = null)
        {

            MarkdownTableBuilder<T> builder = new(tabularData, mdOptions);
            builder.Create(path);
        }

        public static IEnumerable<string> ToMarkdownTable<T>(this TabularData<T> tabularData, MarkdownTableOptions? mdOptions = null)
        {
            MarkdownTableBuilder<T> builder = new(tabularData, mdOptions);
            return builder.Create();
        }

        public static void ToMarkdownTable<T>(this TabularData<T> tabularData, Stream stream, MarkdownTableOptions? mdOptions = null)
        {
            MarkdownTableBuilder<T> builder = new(tabularData, mdOptions);
            builder.WriteStream(stream);
        }
        #endregion
    }
}
