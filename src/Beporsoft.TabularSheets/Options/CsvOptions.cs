using System;
using System.Text;

namespace Beporsoft.TabularSheets.Options
{
    /// <summary>
    /// Configurations for generate CSV files from Tabular sheet.
    /// </summary>
    public class CsvOptions
    {
        /// <summary>
        /// The separator between columns, commonly <see cref="CommaSeparator"/>, <see cref="SemicolonSeparator"/>
        /// <para>
        ///     A special attention deserves to avoid conflicts with culture specific numbering separator.
        /// </para>
        /// </summary>
        public string Separator { get; set; } = SemicolonSeparator;

        /// <summary>
        /// Character encoding for file. Default is UTF-8.
        /// </summary>
        public Encoding Encoding { get; set; } = Encoding.UTF8;

        /// <summary>
        /// Represent the separator ",".
        /// </summary>
        public const string CommaSeparator = ",";

        /// <summary>
        /// Represent the separator ";".
        /// </summary>
        public const string SemicolonSeparator = ";";


    }
}
