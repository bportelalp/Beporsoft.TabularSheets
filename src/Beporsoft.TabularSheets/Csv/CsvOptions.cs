using Beporsoft.TabularSheets.Options;
using System;
using System.Text;

namespace Beporsoft.TabularSheets.Csv
{
    public class CsvOptions
    {
        /// <summary>
        /// The separator between columns. Use <see cref="CommaSeparator"/>, <see cref="SemicolonSeparator"/>
        /// or a custom one.
        /// </summary>
        public string Separator { get; set; } = SemicolonSeparator;

        /// <summary>
        /// The default date time formatting.
        /// </summary>
        public string DateTimeFormat { get; set; } = SheetOptions.DefaultDateTimeFormat;

        public Encoding Encoding { get; set; } = Encoding.GetEncoding("latin1");

        /// <summary>
        /// Represent the separator ",".
        /// 
        /// </summary>
        public const string CommaSeparator = ",";

        /// <summary>
        /// Represent the separator ";".
        /// </summary>
        public const string SemicolonSeparator = ";";


    }
}
