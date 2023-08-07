using System.IO.Compression;

namespace Beporsoft.TabularSheets.Tools.OpenXmlDecompressor
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string currentDirectory = Directory.GetCurrentDirectory(); //"C:\\Users\\bport\\DevOps\\repos\\Beporsoft.TabularSheets\\src\\Beporsoft.TabularSheets.Test\\Results";
            IEnumerable<string> files = Directory.GetFiles(currentDirectory);
            Console.WriteLine("Finding files on {0}", currentDirectory);

            List<string> openXmlFiles = new List<string>();
            foreach (string file in files)
            {
                if(Path.GetExtension(file) == ".xlsx")
                    openXmlFiles.Add(file);
            }

            Console.WriteLine("\n\rTotal Open Xml files: {0}", openXmlFiles.Count);

            foreach (string file in openXmlFiles)
            {
                string directoryEquivalent = Path.GetFileNameWithoutExtension(file);
                string zipEquivalent = Path.ChangeExtension(file, ".zip");
                if (Directory.Exists(directoryEquivalent))
                {
                    //Console.WriteLine("Deleting directory {0}", directoryEquivalent);
                    Directory.Delete(directoryEquivalent, true);
                }
                File.Delete(zipEquivalent);
                File.Copy(file, zipEquivalent);
                ZipFile.ExtractToDirectory(zipEquivalent, directoryEquivalent);
                File.Delete(zipEquivalent);
                Console.WriteLine("Processed file: {0}", Path.GetFileName(file));
            }

            Console.WriteLine("\n\rFinished! Press enter to close this window");
            Console.ReadKey();
        }
    }
}