using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Samples.RestCountries
{

    public class Country
    {
        public Name Name { get; set; } = null!;
        public List<string> Capital { get; set; } = null!;
        public Dictionary<string, Currency> Currencies { get; set; } = null!;
        public string Region { get; set; } = null!;
        public Dictionary<string, string> Languages { get; set; } = null!;
        public int Population { get; set; }
    }

    public class Name
    {
        public string Common { get; set; } = null!;
        public string Official { get; set; } = null!;
    }

    public class Currency
    {
        public string Name { get; set; } = null!;
        public string Symbol { get; set; } = null!;
    }

}
