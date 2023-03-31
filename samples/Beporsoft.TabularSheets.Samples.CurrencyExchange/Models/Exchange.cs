using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Samples.CurrencyExchange.Models
{
    /// <summary>
    /// Represent the conversion between one <see cref="BaseCurrencyCode"/> to <see cref="TargetCurrencyCode"/>
    /// </summary>
    internal class Exchange
    {
        internal string BaseCurrencyCode { get; set; }
        internal string TargetCurrencyCode { get; set; }
        internal double Conversion { get; set; }
    }
}
