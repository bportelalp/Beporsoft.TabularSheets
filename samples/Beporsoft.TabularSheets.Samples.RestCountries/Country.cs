using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

#pragma warning disable CS8618 // Un campo que no acepta valores NULL debe contener un valor distinto de NULL al salir del constructor. Considere la posibilidad de declararlo como que admite un valor NULL.

namespace Beporsoft.TabularSheets.Samples.RestCountries
{

    public class Country
    {
        public Name Name { get; set; }
        public List<string> Capital { get; set; }
        public Dictionary<string, Currency> Currencies { get; set; }
        public string Region { get; set; }
        public Dictionary<string, string> Languages { get; set; }
        public int Population { get; set; }
    }

    public class Name
    {
        public string Common { get; set; }
        public string Official { get; set; }
    }

    public class Currency
    {
        public string Name { get; set; }
        public string Symbol { get; set; }
    }

}
