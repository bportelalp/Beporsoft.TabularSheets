using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text;
using System.Threading.Tasks;
using Beporsoft.TabularSheets.Builders.Interfaces;

namespace Beporsoft.TabularSheets.Builders.Shared
{
    /// <summary>
    /// Represent a collection of <see cref="IIndexedSetup"/>. This is useful to handle collections of nodes
    /// which are indexed and asociated with the respective <see cref="Cell"/> properties.<br/> The collection can register
    /// items that agree the <see cref="{T}"/> constraints to register it only if it hasn't been added yet. Then return
    /// the respective index which can be asociated to <see cref="Cell"/>.<br/> The collection only allow add items, because
    /// removing then can cause unexpected behavior if cells are linked to the setups.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    internal class IndexedSetupCollection<T> where T : Setup, IEquatable<T>, IIndexedSetup
    {
        private readonly List<T> _items = new();


        /// <summary>
        /// Register a new <paramref name="setup"/> f there isn't a equivalent one and return the index to the collection.<br/>
        /// To the respective <see cref="Setup"/> assign its index. Be carefull and avoid modify it.
        /// </summary>
        /// <param name="setup"></param>
        /// <returns></returns>
        public int Register(T setup)
        {
            if (!_items.Contains(setup))
            {
                int index = _items.Count;
                setup.SetIndex(index);
                _items.Add(setup);
            }
            var registerEqual = _items.Single(i => i.Equals(setup));
            return registerEqual.Index;
        }

        /// <summary>
        /// Build the OpenXml container which contains the items of this collection
        /// </summary>
        /// <typeparam name="TContainer">The type which represent parent container</typeparam>
        /// <returns></returns>
        internal TContainer BuildContainer<TContainer>() where TContainer : OpenXmlElement, new()
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
