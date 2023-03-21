using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using System.Text;
using System.Threading.Tasks;
using Beporsoft.TabularSheets.Builders.Interfaces;

namespace Beporsoft.TabularSheets.Builders.StyleBuilders
{
    /// <summary>
    /// Represent a collection of <see cref="IIndexableStyle"/>
    /// </summary>
    /// <typeparam name="T"></typeparam>
    internal class IndexableCollection<T> where T : IEquatable<T>, IIndexableStyle
    {
        private readonly List<T> _items = new();
        public int Register(T style)
        {
            if (_items.Contains(style))
            {
                var founded = _items.Single(i => i.Equals(style));
                return _items.IndexOf(founded);
            }
            else
            {
                _items.Add(style);
                return _items.IndexOf(style);
            }
        }

        public IEnumerable<T> GetRegisteredItems()
        {
            return _items;
        }

        public IEnumerable<OpenXmlElement> GetChildOpenXmlElements()
        {
            foreach (var item in _items)
            {
                yield return item.Build();
            }
        }
    }
}
