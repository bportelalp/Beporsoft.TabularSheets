using Beporsoft.TabularSheets.Builders.Interfaces;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Beporsoft.TabularSheets.Builders.StyleBuilders.Adapters
{
    /// <summary>
    /// Base class to implement <see cref="ISetupCollection{TSetup}"/>
    /// </summary>
    /// <typeparam name="TSetup"></typeparam>
    internal class IndexedSetupCollection<TSetup> : ISetupCollection<TSetup> where TSetup : Setup, IEquatable<TSetup>, IIndexedSetup
    {
        private readonly HashSet<TSetup> _items = new();

        public int Count => _items.Count;

        public TSetup? this[int index]
        {
            get
            {
                if (index < Count)
                    return _items.First(i => i.Index == index);
                else
                    return null;
            }
        }

        public int Register(TSetup setup)
        {
            if (!_items.Contains(setup))
            {
                int index = _items.Count;
                setup.SetIndex(index);
                _items.Add(setup);
                return index;
            }
            else
            {
                var registerEqual = _items.First(i => i.Equals(setup));
                setup.SetIndex(registerEqual.Index);
                return registerEqual.Index;
            }
        }

        public TContainer BuildContainer<TContainer>() where TContainer : OpenXmlElement, new()
        {
            var container = new TContainer();
            var builtItems = _items
                .OrderBy(i => i.Index)
                .Select(i => i.Build())
                .ToList();
            container.Append(builtItems);
            return container;
        }

        public IEnumerable<TSetup> GetRegisteredItems()
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
