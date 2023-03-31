using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Samples.CurrencyExchange.DTO
{
    internal class CurrencyExchangeHistoricalDTO
    {
        public float Amount { get; set; }
        public string Base { get; set; }
        public DateTime Date { get; set; }
        public RatesHistorical Rates { get; set; }
    }

    public class RatesHistorical : Dictionary<DateTime, Rates>
    {
    }
}
