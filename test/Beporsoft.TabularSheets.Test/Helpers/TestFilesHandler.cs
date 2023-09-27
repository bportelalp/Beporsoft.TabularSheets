using DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Test.Helpers
{
    internal class TestFilesHandler
    {
        private static readonly string _baseDir = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory)?.Parent?.Parent?.Parent?.FullName!;
        private string[] _subfolders = Array.Empty<string>();

        public TestFilesHandler() : this(Array.Empty<string>())
        {
        }

        public TestFilesHandler(params string[] subfolders)
        {
            _subfolders = subfolders;
            BasePath = BuildDirectory(subfolders);
            if (!Directory.Exists(BasePath))
            {
                Directory.CreateDirectory(BasePath);
            }
        }

        public string BasePath { get; }

        internal string BuildPath(string fileName) => $"{BasePath}\\{fileName}";

        internal void ClearFiles()
        {
            if (_subfolders.Length > 0 && Directory.Exists(BasePath))
            {
                Directory.Delete(BasePath, true);
            }
            else
            {
                // delete files but neither .gitignore nor the directory
                string[] files = Directory.GetFiles(BasePath);
                foreach (var item in files)
                {
                   if(Path.GetFileName(item) != ".gitignore")
                    {
                        File.Delete(item);
                    }
                }
            }
        }

        internal void ClearFiles(params string[] subfolders)
        {
            string folder = BuildDirectory(subfolders);
            if (Directory.Exists(folder))
            {
                Directory.Delete(folder, true);
            }
        }

        private static string BuildDirectory(params string[] subfolders)
        {
            string folder = "Results";
            foreach (var subfolder in subfolders)
            {
                folder += "\\" + subfolder;
            }
            return $"{_baseDir}\\{folder}";
        }
    }
}
