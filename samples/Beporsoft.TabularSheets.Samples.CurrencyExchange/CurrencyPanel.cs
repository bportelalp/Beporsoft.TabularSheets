using Beporsoft.TabularSheets.Samples.CurrencyExchange.DTO;
using Beporsoft.TabularSheets.Samples.CurrencyExchange.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Json;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Beporsoft.TabularSheets.Samples.CurrencyExchange
{
    public partial class CurrencyPanel : Form
    {
        private readonly HttpClient _apiClient;

        public CurrencyPanel()
        {
            InitializeComponent();
            this._apiClient = new HttpClient { BaseAddress = new Uri("https://api.frankfurter.app/") };
            Configure();
            _ = LoadSelectors();
        }

        #region Building Sheet
        private void BuildExcel(List<ExchangeRecord> records)
        {
            List<string> targetCurrencies = GetTargetCurrencies();

            TabularSheet<ExchangeRecord> table = new TabularSheet<ExchangeRecord>();
            table.AddRange(records);
            table.AddColumn(t => t.Date).SetTitle("Date");
            table.AddColumn(t => t.BaseCurrency).SetTitle("Base Currency");
            // Build a column for each possible currency
            foreach (var currency in targetCurrencies)
            {
                table.AddColumn(t => t.Exchanges.SingleOrDefault(ex => ex.TargetCurrencyCode == currency)?.Conversion)
                    .SetTitle(currency);
            }

            // Add some style
            table.DefaultStyle.DateTimeFormat = CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern;
            table.DefaultStyle.Border.SetAll(CellStyling.BorderStyle.BorderType.Thin);
            table.DefaultStyle.Font.Font = "Calibri";
            table.HeaderStyle.Font.Color = Color.White;
            table.HeaderStyle.Fill.BackgroundColor = Color.Black;


            string fileName = GetPathSave();
            if (fileName != null)
            {
                table.Create(fileName);
                MessageBox.Show("File created");
            }
        }
        #endregion

        #region HandleUIControls
        private void Configure()
        {
            _dataCurrencies.Columns.Clear();
            _dataCurrencies.Columns.Add("Key", "Currency");
            _dataCurrencies.Columns.Add("Value", "Equivalence");
            _dateTimeFrom.Value = DateTime.Today.AddDays(-30);
            _dateTimeTo.Value = DateTime.Today;
            _ = LoadSelectors();
        }
        private async Task LoadSelectors()
        {
            try
            {
                var currencies = await GetCurrencies();

                Currency defaultOrigin = currencies.First();
                if (currencies.Any(c => c.CurrencyCode == "EUR"))
                {
                    defaultOrigin = currencies.Single(c => c.CurrencyCode == "EUR");
                }

                _comboOriginCurrency.Items.Clear();
                _listDestinationCurrencies.Items.Clear();
                _listDestinationCurrencies.CheckOnClick = true;
                foreach (var pair in currencies)
                {
                    _comboOriginCurrency.Items.Add(pair);
                    int index = _listDestinationCurrencies.Items.Add(pair);
                    if (index != -1)
                        _listDestinationCurrencies.SetItemChecked(index, true);
                }
                _comboOriginCurrency.SelectedItem = defaultOrigin;
            }
            catch (Exception)
            {
                MessageBox.Show("Data of currencies cannot be loaded", "Oops!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private List<string> GetTargetCurrencies()
        {
            var targets = _listDestinationCurrencies.CheckedItems.Cast<Currency>().Select(x => x.CurrencyCode).ToList();
            var selected = _comboOriginCurrency.SelectedItem as Currency;
            return targets.Where(c => c != selected.CurrencyCode).ToList();
        }
        private void FillGrid(ExchangeRecord currencyRate)
        {
            _dataCurrencies.Rows.Clear();
            foreach (var rate in currencyRate.Exchanges)
            {
                _dataCurrencies.Rows.Add(rate.TargetCurrencyCode, rate.Conversion);
            }
        }
        #endregion

        #region API Requests
        private async Task<List<Currency>> GetCurrencies()
        {
            var response = await _apiClient.GetAsync("currencies");
            var currencies = await response.Content.ReadFromJsonAsync<Dictionary<string, string>>();
            List<Currency> list = new List<Currency>();
            foreach (var pair in currencies)
            {
                list.Add(new Currency()
                {
                    CurrencyCode = pair.Key,
                    CurrencyName = pair.Value,
                });
            }
            return list;
        }

        private async Task<ExchangeRecord> GetExchangeLatest(string currency, params string[] targetCurrencies)
        {
            string endpoint = $"latest?from={currency}";
            if (targetCurrencies.Length > 0)
            {
                endpoint += "&to=";
                foreach (var curr in targetCurrencies)
                    endpoint += curr + ",";
            }
            var response = await _apiClient.GetAsync(endpoint);
            var content = await response.Content.ReadFromJsonAsync<CurrencyExchangeLatestDTO>();
            ExchangeRecord record = new ExchangeRecord()
            {
                BaseCurrency = content.Base,
                Date = content.Date,
            };
            foreach (var rateRecord in content.Rates)
            {
                record.Exchanges.Add(new Exchange()
                {
                    BaseCurrencyCode = content.Base,
                    TargetCurrencyCode = rateRecord.Key,
                    Conversion = rateRecord.Value / content.Amount
                });
            }
            return record;
        }

        private async Task<List<ExchangeRecord>> GetExchangeHistory(string currencyCode, DateTime from, DateTime to, params string[] targetCurrencies)
        {
            const string format = "yyyy-MM-dd";
            var endpoint = $"{from.ToString(format)}..{to.ToString(format)}?from={currencyCode}";
            if (targetCurrencies.Length > 0)
            {
                endpoint += "&to=";
                foreach (var curr in targetCurrencies)
                    endpoint += curr + ",";
            }
            var response = await _apiClient.GetAsync(endpoint);
            var historical = await response.Content.ReadFromJsonAsync<CurrencyExchangeHistoricalDTO>();
            var list = new List<ExchangeRecord>();
            foreach (var historicalStamp in historical.Rates)
            {
                ExchangeRecord record = new ExchangeRecord()
                {
                    BaseCurrency = historical.Base,
                    Date = historicalStamp.Key,
                };
                foreach (var exch in historicalStamp.Value)
                {
                    record.Exchanges.Add(new Exchange()
                    {
                        BaseCurrencyCode = historical.Base,
                        TargetCurrencyCode = exch.Key,
                        Conversion = exch.Value / historical.Amount,
                    });
                }
                list.Add(record);

            }
            return list;
        }
        #endregion

        #region UIEvents
        private async void OnSearch(object sender, EventArgs e)
        {
            Currency selected = _comboOriginCurrency.SelectedItem as Currency;
            string currencyCode = selected.CurrencyCode;
            List<string> to = GetTargetCurrencies();
            var latest = await GetExchangeLatest(currencyCode, to.ToArray());
            FillGrid(latest);
        }

        private async void ExportSheetHistory(object sender, EventArgs e)
        {
            DateTime from = _dateTimeFrom.Value;
            DateTime to = _dateTimeTo.Value;
            Currency selected = _comboOriginCurrency.SelectedItem as Currency;
            string currencyCode = selected.CurrencyCode;
            if (from > to)
            {
                MessageBox.Show("The input dates are incoherent!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            List<string> targetCurrencies = GetTargetCurrencies();
            var result = await GetExchangeHistory(currencyCode, from, to, targetCurrencies.ToArray());
            BuildExcel(result);
        }

        private string GetPathSave()
        {
            _saveFileDialog.Filter = "Excel files|.xlsx";
            _saveFileDialog.FileName = "Currency.xlsx";

            if (_saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                return _saveFileDialog.FileName;
            }
            return null;

        }

        private void BtnSelectAll(object sender, EventArgs e)
        {
            for (int i = 0; i < _listDestinationCurrencies.Items.Count; i++)
            {
                _listDestinationCurrencies.SetItemChecked(i, true);
            }
        }

        private void BtnUnselectAll(object sender, EventArgs e)
        {
            for (int i = 0; i < _listDestinationCurrencies.Items.Count; i++)
            {
                _listDestinationCurrencies.SetItemChecked(i, false);
            }
        }
        #endregion


    }
}
