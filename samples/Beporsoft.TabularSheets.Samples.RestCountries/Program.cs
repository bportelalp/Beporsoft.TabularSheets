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
            Console.WriteLine($"Retrieving information about {region}");
            Uri uri = CreateUri(region);
            HttpClient client = new HttpClient();
            client.BaseAddress = uri;

            HttpResponseMessage response = await client.GetAsync(uri);

            List<Country>? countries = await response.Content.ReadFromJsonAsync<List<Country>>();
            Console.WriteLine($"Obtained information about {countries?.Count} countries");

            if (countries is not null)
            {
                // Create tabularsheet
                TabularSheet<Country> sheet = new TabularSheet<Country>();
                sheet.AddRange(countries.OrderBy(c => c.Name.Common));
                sheet.SetSheetTitle("List of european countries");

                // Configure columns
                Console.WriteLine($"Creating columns");
                sheet.AddColumn("Common name", c => c.Name.Common);
                sheet.AddColumn("Official name", c => c.Name.Official);
                sheet.AddColumn("Region", c => c.Region);
                sheet.AddColumn("Capital", c => c.Capital.FirstOrDefault()!);
                sheet.AddColumn("Languages", c => string.Join("; ", c.Languages.Values));
                sheet.AddColumn("Population", c => c.Population);
                sheet.AddColumn("Currencies", c => string.Join("; ", c.Currencies.Values.Select(v => $"{v.Name} ({v.Symbol})")));


                // Add some style
                Console.WriteLine($"Adding some style");
                sheet.HeaderStyle.Fill.BackgroundColor = Color.Black;
                sheet.HeaderStyle.Font.FontColor = Color.White;
                sheet.DefaultStyle.Border.Top = Styling.BorderStyle.BorderType.Thin;
                sheet.DefaultStyle.Border.Bottom = Styling.BorderStyle.BorderType.Thin;

                // Export
                Console.WriteLine($"Creating file");
                string path = PrepareDirectory($"{region}-countries.xlsx");
                sheet.Create(path);
                Console.WriteLine($"Done!");
            }
            else
            {
                Console.WriteLine("Error retrieving results from REST Countries");
            }
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
                var input = Console.ReadLine();
                bool converted = int.TryParse(input, out int result);
                if (converted && result > 0 && result <= _regions.Count)
                    region = _regions[result - 1];
            }
            return region;
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
            string projectDirectory = Directory.GetParent(Environment.CurrentDirectory)!.Parent!.Parent!.FullName;
            string resultDirectory = Path.Combine(projectDirectory, "Results");
            Directory.CreateDirectory(resultDirectory);
            string resultPath = Path.Combine(resultDirectory, filename);
            return resultPath;
        }
    }
}