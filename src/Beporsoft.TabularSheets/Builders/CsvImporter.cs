using DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Builders
{
    public class CsvImporter
    {
        public List<T> FromCsv<T>(string path) where T : class, new()
        {
            List<T> result = new List<T>();
            List<string> lines = new List<string>();
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read))
            using (var sr = new StreamReader(fs))
            {
                while(!sr.EndOfStream)
                {
                    lines.Add(sr.ReadLine());
                }
            }

            string[] cols = lines[0].Split(';');
            List<string> data = lines.Skip(1).ToList();
            foreach (var line in data)
            {
                T row = new();
                var values = line.Split(';');
                for (int i = 0; i <= cols.Length; i++)
                {
                    PropertyInfo prop = typeof(T).GetProperty(cols[i]);
                    object value = Convert.ChangeType(values[i], prop.PropertyType);
                    prop.SetValue(row, value, null);
                }
                result.Add(row);
            }
            return result;
        }
    }
}
