using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
                    throw new FileLoadException($"The path provided does not aim a {targetExtension} file extension", path);
                else
                    return path;
            }
            else
            {
                return $"{path}{targetExtension.First()}";
            }
        }
    }

    internal struct SpreadsheetsFileExtension
    {
        public static string Excel2007_365 = ".xlsx";
        public static string MIME_Excel2007_365 = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        public static string Excel97_2003 = ".xls";
        public static string MIME_Excel97_2003 = "application/vnd.ms-excel";
        public static string[] AllowedExtensions = new string[] { Excel2007_365, Excel97_2003 };
        public static string[] AllowedMimeTypes = new string[] { MIME_Excel2007_365, Excel97_2003 };
    }

    internal struct CsvFileExtension
    {
        public static string Csv = ".csv";
        public static string MIME_Csv = "text/csv";
    }
}
