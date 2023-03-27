using Beporsoft.TabularSheets.Builders.StyleBuilders.SetupCollections;
using DocumentFormat.OpenXml;

namespace Beporsoft.TabularSheets.Builders.StyleBuilders
{
    internal abstract class Setup
    {
        public int Index { get; private set; }

        public abstract OpenXmlElement Build();

        /// <summary>
        /// Establish the index of the setup inside its parent <see cref="IndexedSetupCollection{T}"/>
        /// </summary>
        /// <param name="index"></param>
        internal void SetIndex(int index) => Index = index;
    }
}
