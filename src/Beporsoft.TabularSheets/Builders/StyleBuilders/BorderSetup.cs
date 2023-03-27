using Beporsoft.TabularSheets.Builders.Interfaces;
using Beporsoft.TabularSheets.Style;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace Beporsoft.TabularSheets.Builders.StyleBuilders
{
    [DebuggerDisplay("Id={Index}")]
    internal class BorderSetup : Setup, IIndexedSetup, IEquatable<BorderSetup?>
    {
        internal BorderSetup(BorderStyle borderStyle)
        {
            BorderStyle = borderStyle;
        }

        public BorderStyle BorderStyle { get; set; }

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
                EqualityComparer<BorderStyle>.Default.Equals(BorderStyle, other.BorderStyle);
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(BorderStyle);
        }
    }
}
