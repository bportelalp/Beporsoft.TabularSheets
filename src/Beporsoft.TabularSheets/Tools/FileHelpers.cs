using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Tools
{
    internal static class FileHelpers
    {
        internal static string VerifyPath(string path, string targetExtension)
        {
            bool hasExtension = Path.HasExtension(path);
            if (hasExtension)
            {
                string extension = Path.GetExtension(path).Trim();
                if (extension != targetExtension)
                    throw new FileLoadException($"The path provided does not aim a {targetExtension} file extension", path);
                else
                    return path;
            }
            else
            {
                return $"{path}{targetExtension}";
            }
        }
    }
}
