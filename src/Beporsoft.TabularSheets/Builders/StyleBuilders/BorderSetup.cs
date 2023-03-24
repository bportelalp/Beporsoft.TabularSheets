using Beporsoft.TabularSheets.Builders.Interfaces;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Builders.StyleBuilders
{
    internal class BorderSetup : IStyleSetup
    {
        internal BorderSetup()
        {
            
        }
        public int Index { get; set; }

        public OpenXmlElement Build()
        {
            throw new NotImplementedException();
        }
    }
}
