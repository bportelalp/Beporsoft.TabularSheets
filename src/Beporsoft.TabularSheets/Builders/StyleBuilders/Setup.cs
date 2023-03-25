using Beporsoft.TabularSheets.Builders.Interfaces;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Builders.StyleBuilders
{
    internal abstract class Setup : IStyleSetup
    {
        public int Index { get; set; }

        public abstract OpenXmlElement Build();
    }
}
