using Beporsoft.TabularSheets.Builders.StyleBuilders;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace Beporsoft.TabularSheets.Builders.Interfaces
{
    /// <summary>
    /// Represent a collection of <see cref="IIndexedSetup"/>. This is useful to handle collections of nodes
    /// which are indexed and asociated with the respective <see cref="Cell"/> properties.<br/> The collection can register
    /// items that agree the <typeparamref name="TSetup"/> constraints to register it only if it hasn't been added yet. Then return
    /// the respective index which can be asociated to <see cref="Cell"/>.<br/> The collection only allow add items, because
    /// removing then can cause unexpected behavior if cells are linked to the setups.
    /// </summary>
    /// <typeparam name="TSetup"></typeparam>
    internal interface ISetupCollection<TSetup> where TSetup : Setup, IEquatable<TSetup>, IIndexedSetup
    {

        /// <summary>
        /// Amount of <typeparamref name="TSetup"/> items contained in this collection
        /// </summary>
        internal int Count { get; }

        /// <summary>
        /// Register a new <paramref name="setup"/> if there isn't a equivalent one and return it index inside the collection.<br/>
        /// Be carefull and avoid modify it.
        /// </summary>
        /// <param name="setup"></param>
        /// <returns>The index assigned to <paramref name="setup"/> inside the collection</returns>
        internal int Register(TSetup setup);

        /// <summary>
        /// Build the OpenXml container which contains the items of this collection
        /// </summary>
        /// <typeparam name="TContainer">The type which represent parent container</typeparam>
        /// <returns>The <see cref="OpenXmlElement"/> which acts as container of <typeparamref name="TSetup"/> items</returns>
        internal TContainer BuildContainer<TContainer>() where TContainer : OpenXmlElement, new();
    }
}
