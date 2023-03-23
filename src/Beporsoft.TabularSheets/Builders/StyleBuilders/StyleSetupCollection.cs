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
    /// Represent a collection of <see cref="IStyleSetup"/>
    /// </summary>
    /// <typeparam name="T"></typeparam>
    internal class StyleSetupCollection<T> where T : IEquatable<T>, IStyleSetup
    {
        private readonly List<T> _items = new();
        public int Register(T style)
        {
            if (!_items.Contains(style))
            {
                int index = _items.Count;
                style.Index = index;
                _items.Add(style);
            }
            var registerEqual = _items.Single(i => i.Equals(style));
            return registerEqual.Index;
        }

        public TContainer GetContainer<TContainer>() where TContainer : OpenXmlElement, new()
        {
            var container = new TContainer();
            var builtItems = _items.Select(i => i.Build());
            container.Append(builtItems);
            return container;
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
