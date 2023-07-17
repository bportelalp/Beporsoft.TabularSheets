using Beporsoft.TabularSheets.Builders.Interfaces;
using Beporsoft.TabularSheets.Builders.SheetBuilders;
using DocumentFormat.OpenXml;
using System.Collections.Generic;
using System.Linq;

namespace Beporsoft.TabularSheets.Builders.StyleBuilders.SetupCollections
{
    /// <summary>
    /// Specific class to implement ISetupCollection for sharedStrings. Improve the preformance using a internal dictionary being the
    /// string the key. Due to dictionary is a lookup, the checking of contained string performs better than a list, which uses iteration.
    /// For large sets of data, this shorten the search time of coincidences
    /// </summary>
    internal class SharedStringSetupCollection : ISetupCollection<SharedStringSetup>
    {
        private readonly Dictionary<string, SharedStringSetup> _items = new();

        public int Count => _items.Count;

        public int Register(SharedStringSetup setup)
        {
            if (!_items.ContainsKey(setup.Text))
            {
                int index = _items.Count;
                setup.SetIndex(index);
                _items[setup.Text] = setup;
            }
            var registerEqual = _items[setup.Text];
            setup.SetIndex(registerEqual.Index);
            return registerEqual.Index;
        }

        public TContainer BuildContainer<TContainer>() where TContainer : OpenXmlElement, new()
        {
            var container = new TContainer();
            var builtItems = _items.Select(i => i.Value.Build());
            container.Append(builtItems);
            return container;
        }

        public IEnumerable<SharedStringSetup> GetRegisteredItems()
        {
            return _items.Values;
        }

        public IEnumerable<OpenXmlElement> GetChildOpenXmlElements()
        {
            foreach (var item in _items)
            {
                yield return item.Value.Build();
            }
        }
    }
}
