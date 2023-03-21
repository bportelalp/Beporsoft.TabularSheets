using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Builders.StyleBuilders
{
    internal class IndexableFormat
    {
        public IndexableFormat(IndexableFill? fill)
        {
            Fill = fill;
        }

        public IndexableFill? Fill { get; }
    }
}
