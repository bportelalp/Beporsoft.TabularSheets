using Beporsoft.TabularSheets.Builders.Interfaces;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Builders.Shared
{
    internal abstract class Setup : IIndexedSetup
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
