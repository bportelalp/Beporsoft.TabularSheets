namespace Beporsoft.TabularSheets.Options
{
    /// <summary>
    /// Options to configure the markdown table generation process.
    /// </summary>
    public class MarkdownTableOptions
    {
        /// <summary>
        /// Default separator for tables according to Markdown general specification.
        /// </summary>
        public const string Separator = "|";

        /// <summary>
        /// Default configuration.
        /// </summary>
        public static MarkdownTableOptions Default { get; } = new MarkdownTableOptions();

        /// <summary>
        /// Supress titles writing on table. <br/>
        /// Note: Header will be added but titles are skipped. This is due to most Markdown Parsers don't support headerless
        /// tables, but a workaround is to let it empty. See 
        /// <see href="https://stackoverflow.com/questions/17536216/create-a-table-without-a-header-in-markdown">
        ///     Create a table without a header in markdown
        /// </see>
        /// </summary>
        public bool SupressHeaderTitles { get; set; } = false;

        /// <summary>
        /// If <see langword="true"/>, extra spaces to align cells will be removed. Default is <see langword="false"/>.<br/>
        /// Note: For large tables, it is recomended compact mode to avoid the requirement of measure the largest text before
        /// build the table.
        /// </summary>
        public bool CompactMode { get; set; } = false;
    }
}
