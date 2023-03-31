using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Samples.CurrencyExchange.DTO
{
    public class CurrencyExchangeLatestDTO
    {
        public float Amount { get; set; }
        public string Base { get; set; }
        public DateTime Date { get; set; }
        public Rates Rates { get; set; }
    }

    public class Rates : Dictionary<string, double>
    {
    }

}
