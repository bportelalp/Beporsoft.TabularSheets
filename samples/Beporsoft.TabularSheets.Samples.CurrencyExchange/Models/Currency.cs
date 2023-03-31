using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Samples.CurrencyExchange.Models
{
    internal class Currency
    {
        internal string CurrencyCode { get; set; }
        internal string CurrencyName { get; set; }

        public override string ToString()
        {
            return $"{CurrencyName} ({CurrencyCode})";
        }
    }
}
