using Beporsoft.TabularSheets.Builders.Interfaces;
using DocumentFormat.OpenXml;
using System.Collections.Generic;
using System.Linq;

namespace Beporsoft.TabularSheets.Builders.StyleBuilders.SetupCollections
{
    internal class NumberingFormatSetupCollection : ISetupCollection<NumberingFormatSetup>
    {
        private readonly List<NumberingFormatSetup> _items = new();

        public int Count => _items.Where(i => i.Index >= NumberingFormatSetup.StartIndexNotBuiltin).Count();
        public int Register(NumberingFormatSetup setup)
        {
            if (!_items.Contains(setup))
            {
                if (PredefinedFormats.ContainsValue(setup.Pattern))
                {
                    var pair = PredefinedFormats.Single(p => p.Value == setup.Pattern);
                    setup.SetIndex(pair.Key);
                    _items.Add(setup);
                    return pair.Key;
                }
                else
                {
                    int maxIndex = NumberingFormatSetup.StartIndexNotBuiltin;
                    if (_items.Any())
                    {
                        int maxIndexItems = _items.Max(i => i.Index);
                        if (maxIndexItems >= maxIndex)
                            maxIndex = maxIndexItems + 1;
                    }
                    setup.SetIndex(maxIndex);
                    _items.Add(setup);
                    return maxIndex;
                }
            }
            var registerEqual = _items.Single(i => i.Equals(setup));
            setup.SetIndex(registerEqual.Index);
            return registerEqual.Index;
        }
        public TContainer BuildContainer<TContainer>() where TContainer : OpenXmlElement, new()
        {
            var container = new TContainer();
            var notBuiltInItems = _items.Where(i => i.Index >= NumberingFormatSetup.StartIndexNotBuiltin);
            var builtItems = notBuiltInItems.Select(i => i.Build());
            container.Append(builtItems);
            return container;
        }



        internal static readonly Dictionary<int, string> PredefinedFormats = new Dictionary<int, string>()
        {
            { 0 , "General"},
            { 1 , "0"},
            { 2 , "0.00"},
            { 3 , "#,##0"},
            { 4 , "#,##0.00"},
            { 5 , "$#,##0},\\-$#,##0"},
            { 6 , "$#,##0},[Red]\\-$#,##0"},
            { 7 , "$#,##0.00},\\-$#,##0.00"},
            { 8 , "$#,##0.00},[Red]\\-$#,##0.00"},
            { 9 , "0%"},
            { 10 , "0.00%"},
            { 11 , "0.00E+00"},
            { 12 , "# ?/?"},
            { 13 , "# ??/??"},
            { 14 , "mm-dd-yy"},
            { 15 , "d-mmm-yy"},
            { 16 , "d-mmm"},
            { 17 , "mmm-yy"},
            { 18 , "h:mm AM/PM"},
            { 19 , "h:mm:ss AM/PM"},
            { 20 , "h:mm"},
            { 21 , "h:mm:ss"},
            { 22 , "m/d/yy h:mm"},
            { 37 , "#,##0 },(#,##0)"},
            { 38 , "#,##0 },[Red](#,##0)"},
            { 39 , "#,##0.00},(#,##0.00)"},
            { 40 , "#,##0.00},[Red](#,##0.00)"},
            { 44 , "_(\"$\"* #,##0.00_)},_(\"$\"* \\(#,##0.00\\)},_(\"$\"* \"-\"??_)},_(@_)"},
            { 45 , "mm:ss"},
            { 46 , "[h]:mm:ss"},
            { 47 , "mmss.0"},
            { 48 , "##0.0E+0"},
            { 49 , "@"},
            { 27 , "[$-404]e/m/d"},
            { 30 , "m/d/yy"},
            { 36 , "[$-404]e/m/d"},
            { 50 , "[$-404]e/m/d"},
            { 57 , "[$-404]e/m/d"},
            { 59 , "t0"},
            { 60 , "t0.00"},
            { 61 , "t#,##0"},
            { 62 , "t#,##0.00"},
            { 67 , "t0%"},
            { 68 , "t0.00%"},
            { 69 , "t# ?/?"},
            { 70 , "t# ??/??"},
        };
    }
}
