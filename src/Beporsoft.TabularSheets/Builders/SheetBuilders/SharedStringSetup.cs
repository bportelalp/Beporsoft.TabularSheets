using Beporsoft.TabularSheets.Builders.Interfaces;
using Beporsoft.TabularSheets.Builders.Shared;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Builders.SheetBuilders
{
    internal class SharedStringSetup : Setup, IEquatable<SharedStringSetup>, IIndexedSetup
    {
        public SharedStringSetup(string text)
        {
            Text = text;
        }

        public string Text { get; set; }

        public override OpenXmlElement Build()
        {
            var sharedString = new SharedStringItem()
            {
                Text = new Text(Text)
            };
            return sharedString;
        }

        public override bool Equals(object? obj)
        {
            return Equals(obj as SharedStringSetup);
        }

        public bool Equals(SharedStringSetup? other)
        { 
            return other is SharedStringSetup setup &&
                   Text == setup.Text;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(Index, Text);
        }
    }
}
