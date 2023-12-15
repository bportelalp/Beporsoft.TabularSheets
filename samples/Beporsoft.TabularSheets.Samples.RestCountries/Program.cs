using Beporsoft.TabularSheets.CellStyling;
using Beporsoft.TabularSheets.Options.ColumnWidth;
using System.Drawing;
using System.Net.Http.Json;

namespace Beporsoft.TabularSheets.Samples.RestCountries
{
    internal class Program
    {
        static List<string> _regions = new List<string>() { "All", "Europe", "America", "Africa", "Asia", "Oceania" };

        static async Task Main(string[] args)
        {
            string region = SelectRegion();
            bool eachRegionSingleSheet = false;
            if (region == "All")
                eachRegionSingleSheet = SelectAllCountriesEachRegionSheet();
            string fileFormat = string.Empty;

            if (!eachRegionSingleSheet)
                fileFormat = SelectFileFormat();

            List<Country>? countries = await GetCountriesAsync(region);

            if (countries is not null)
            {
                if (!eachRegionSingleSheet)
                    CreateSingleCountrySheet(countries, region, fileFormat);
                else
                    CreateMultipleCountrySheet(countries);
            }
            else
            {
                Console.WriteLine("Error retrieving results from REST Countries");
            }
        }

        private static void CreateSingleCountrySheet(List<Country> countries, string region, string fileFormat)
        {
            TabularSheet<Country> sheet = FillTabularSheet(countries, region);

            // Export
            Console.WriteLine($"Creating file");
            string path = PrepareDirectory($"{region}-countries{fileFormat}");
            bool fileFormatNotOk = false;
            if (fileFormat == ".xlsx")
                sheet.Create(path);
            else if (fileFormat == ".csv")
                sheet.ToCsv(path);
            else
                fileFormatNotOk = true;

            if (!fileFormatNotOk)
                Console.WriteLine($"Done! Exported on: {path}");
            else
                Console.WriteLine($"File format {fileFormat} is unsupported");
        }

        private static void CreateMultipleCountrySheet(List<Country> countries)
        {
            var countriesByRegion = countries.GroupBy(country => country.Region);
            TabularBook workbook = new TabularBook();
            foreach (var country in countriesByRegion)
            {
                TabularSheet<Country> sheet = FillTabularSheet(country.ToList(), country.Key);
                workbook.Add(sheet);
            }
            Console.WriteLine($"Pack all on single workbook");

            // Export
            Console.WriteLine($"Creating file");
            string path = PrepareDirectory($"All-countries.xlsx");
            workbook.Create(path);
            Console.WriteLine($"Done! Exported on: {path}");
        }

        private static TabularSheet<Country> FillTabularSheet(List<Country> countries, string region)
        {
            Console.WriteLine($"Fill sheet of region: {region}");
            // Create tabularsheet
            TabularSheet<Country> sheet = new TabularSheet<Country>();
            sheet.AddRange(countries.OrderBy(c => c.Name.Common));
            sheet.SetSheetTitle(region);

            // Configure columns
            sheet.AddColumn("Common name", c => c.Name.Common);
            sheet.AddColumn("Official name", c => c.Name.Official);
            sheet.AddColumn("Region", c => c.Region);
            sheet.AddColumn("Capital", c => c.Capital.FirstOrDefault()!);
            sheet.AddColumn("Languages", c => string.Join("; ", c.Languages.Values));
            sheet.AddColumn("Population", c => c.Population)
                .SetStyle(s => s.NumberingPattern = "#,##0"); // Style with no decimals and thousands separator
            sheet.AddColumn("Currencies", c => string.Join("; ", c.Currencies.Values.Select(v => $"{v.Name} ({v.Symbol})")));

            // Add some style
            sheet.HeaderStyle.Fill.BackgroundColor = Color.DarkOliveGreen;
            sheet.HeaderStyle.Font.Color = Color.White;
            sheet.HeaderStyle.Border.Bottom = BorderStyle.BorderType.Medium;

            sheet.BodyStyle.Border.SetBorderType(BorderStyle.BorderType.Thin, null);
            sheet.BodyStyle.Font.FontName = "Calibri";
            sheet.Options.InheritHeaderStyleFromBody = true;
            sheet.Options.ColumnOptions.Width = new AutoColumnWidth();

            return sheet;
        }


        private static async Task<List<Country>?> GetCountriesAsync(string region)
        {
            Console.WriteLine($"Retrieving information about {region}");
            Uri uri = CreateUri(region);
            HttpClient client = new HttpClient();
            client.BaseAddress = uri;

            HttpResponseMessage response = await client.GetAsync(uri);

            List<Country>? countries = await response.Content.ReadFromJsonAsync<List<Country>>();
            Console.WriteLine($"Obtained information about {countries?.Count} countries");
            return countries;
        }

        private static string SelectRegion()
        {
            string? region = null;
            while (region is null)
            {
                int index = 1;
                Console.WriteLine("Select region");
                foreach (var reg in _regions)
                {
                    Console.WriteLine($"{index}. {reg}");
                    index++;
                }
                Console.WriteLine(Environment.NewLine);
                Console.WriteLine("Type a number and press enter");
                var input = Console.ReadLine();
                bool converted = int.TryParse(input, out int result);
                if (converted && result > 0 && result <= _regions.Count)
                    region = _regions[result - 1];
            }
            return region;
        }

        private static bool SelectAllCountriesEachRegionSheet()
        {
            bool? response = null;
            while (response is null)
            {
                Console.WriteLine("You've selected all regions. ¿Do you want each region on a single sheet on the workbook? (Y/N)");
                var input = Console.ReadLine();
                if (input?.ToUpper() == "Y")
                    response = true;
                else if (input?.ToUpper() == "N")
                    response = false;
            }
            return response.Value;
        }

        private static string SelectFileFormat()
        {
            string[] formats = new string[] { ".xlsx", ".csv" };
            string? fileFormat = null;
            while (fileFormat is null)
            {
                int index = 1;
                Console.WriteLine("Select file format");
                foreach (var format in formats)
                {
                    Console.WriteLine($"{index}. {format}");
                    index++;
                }
                Console.WriteLine(Environment.NewLine);
                Console.WriteLine("Type a number and press enter");
                var input = Console.ReadLine();
                bool converted = int.TryParse(input, out int result);
                if (converted && result > 0 && result <= _regions.Count)
                    fileFormat = formats[result - 1];
            }
            return fileFormat;
        }

        private static Uri CreateUri(string region)
        {
            Uri uri;
            if (region == "All")
                uri = new Uri("https://restcountries.com/v3.1/all?fields=name,capital,population,currencies,region,languages");
            else
                uri = new Uri($"https://restcountries.com/v3.1/region/{region.ToLower()}?fields=name,capital,population,currencies,region,languages");
            return uri;
        }

        private static string PrepareDirectory(string filename)
        {
            string? projectDirectory = Directory.GetParent(Environment.CurrentDirectory)?.Parent?.Parent?.FullName;
            string resultDirectory = string.Empty;
            if (!string.IsNullOrWhiteSpace(projectDirectory))
            {
                resultDirectory = Path.Combine(projectDirectory, "Results");
                Directory.CreateDirectory(resultDirectory);
            }
            string resultPath = Path.Combine(resultDirectory, filename);
            return resultPath;
        }
    }
}