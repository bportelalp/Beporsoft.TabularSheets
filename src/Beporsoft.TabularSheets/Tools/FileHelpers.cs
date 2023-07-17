using System;
using System.IO;
using System.Linq;

namespace Beporsoft.TabularSheets.Tools
{
    internal static class FileHelpers
    {
        internal static string VerifyPath(string path, params string[] targetExtension)
        {
            bool hasExtension = Path.HasExtension(path);
            if (hasExtension)
            {
                string extension = Path.GetExtension(path).Trim();
                if (!targetExtension.Contains(extension))
                    throw new FileLoadException($"The path provided does not aim one of the following file extensions [{string.Join(", ", targetExtension)}]", path);
                else
                    return path;
            }
            else
            {
                return $"{path}{targetExtension.First()}";
            }
        }
    }

    /// <summary>
    /// File extensions allowed for building spreadsheets
    /// </summary>
    public struct SpreadsheetFileExtension
    {
        /// <summary>
        /// File extension for Microsoft Excel spreadsheet documents according to the Office Open XML SpreadsheetML file format.
        /// </summary>
        public static string Excel2007_365 = ".xlsx";
        internal static string Excel97_2003 = ".xls";
        internal static string[] AllowedExtensions = new string[] { Excel2007_365 };
    }

    internal struct SpreadsheetMimeTypes
    {
        public static string Excel2007_365 = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        public static string Excel97_2003 = "application/vnd.ms-excel";
        public static string[] AllowedMimeTypes = new string[] { Excel2007_365, Excel97_2003 };

    }

    internal struct CsvFileExtension
    {
        public static string Csv = ".csv";
        public static string MIME_Csv = "text/csv";
    }
}
