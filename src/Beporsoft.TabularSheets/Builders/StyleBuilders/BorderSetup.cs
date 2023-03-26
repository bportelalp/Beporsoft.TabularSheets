using Beporsoft.TabularSheets.Builders.Interfaces;
using Beporsoft.TabularSheets.Builders.Shared;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Builders.StyleBuilders
{
    internal class BorderSetup : Setup, IIndexedSetup, IEquatable<BorderSetup?>
    {
        internal BorderSetup()
        {
            
        }

        public override OpenXmlElement Build()
        {
            throw new NotImplementedException();
        }

        public override bool Equals(object? obj)
        {
            return Equals(obj as BorderSetup);
        }

        public bool Equals(BorderSetup? other)
        {
            return other is not null &&
                   Index == other.Index;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(Index);
        }
    }
}
