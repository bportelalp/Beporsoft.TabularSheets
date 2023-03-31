using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Samples.CurrencyExchange.Models
{
    internal class ExchangeRecord
    {
        public DateTime Date { get; set; }   
        public string BaseCurrency { get; set; }
        public List<Exchange> Exchanges { get; set; } = new List<Exchange>();
    }
}
