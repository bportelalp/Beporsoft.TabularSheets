namespace Beporsoft.TabularSheets.Options
{
    public class MarkdownTableOptions
    {
        /// <summary>
        /// Default separator for tables according to Markdown general specification.
        /// </summary>
        public const string Separator = "|";

        public static MarkdownTableOptions Default { get; } = new MarkdownTableOptions();

        /// <summary>
        /// Supress titles writing on table. <br/>
        /// Note: Header will be added but titles are skipped. This is due to most Markdown parses don't support headerless
        /// tables, but a workaround is to let it empty. See 
        /// <see href="https://stackoverflow.com/questions/17536216/create-a-table-without-a-header-in-markdown">
        ///     Create a table without a header in markdown
        /// </see>
        /// </summary>
        public bool SupressHeaderTitles { get; set; } = false;
        public bool PrettyPrint { get; set; } = false;
    }
}
