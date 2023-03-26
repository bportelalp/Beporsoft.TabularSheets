using Beporsoft.TabularSheets.Builders.StyleBuilders;
using DocumentFormat.OpenXml;
using System;

namespace Beporsoft.TabularSheets.Builders.Interfaces
{
    /// <summary>
    /// Represent a collection of <see cref="IIndexedSetup"/>. This is useful to handle collections of nodes
    /// which are indexed and asociated with the respective <see cref="Cell"/> properties.<br/> The collection can register
    /// items that agree the <see cref="{T}"/> constraints to register it only if it hasn't been added yet. Then return
    /// the respective index which can be asociated to <see cref="Cell"/>.<br/> The collection only allow add items, because
    /// removing then can cause unexpected behavior if cells are linked to the setups.
    /// </summary>
    /// <typeparam name="TSetup"></typeparam>
    internal interface ISetupCollection<TSetup> where TSetup : Setup, IEquatable<TSetup>, IIndexedSetup
    {
        public int Count { get; }

        /// <summary>
        /// Register a new <paramref name="setup"/> f there isn't a equivalent one and return the index to the collection.<br/>
        /// To the respective <see cref="Setup"/> assign its index. Be carefull and avoid modify it.
        /// </summary>
        /// <param name="setup"></param>
        /// <returns></returns>
        public int Register(TSetup setup);

        /// <summary>
        /// Build the OpenXml container which contains the items of this collection
        /// </summary>
        /// <typeparam name="TContainer">The type which represent parent container</typeparam>
        /// <returns></returns>
        public TContainer BuildContainer<TContainer>() where TContainer : OpenXmlElement, new();
    }
}
