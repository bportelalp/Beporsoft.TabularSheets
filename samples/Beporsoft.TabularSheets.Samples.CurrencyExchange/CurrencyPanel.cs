using Beporsoft.TabularSheets.CellStyling;
using Beporsoft.TabularSheets.Samples.CurrencyExchange.DTO;
using Beporsoft.TabularSheets.Samples.CurrencyExchange.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.CompilerServices;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Beporsoft.TabularSheets.Samples.CurrencyExchange
{
    public partial class CurrencyPanel : Form
    {
        private readonly HttpClient _apiClient;
        private bool _applyFontHeader = false;
        private bool _applyFontBody = false;
        public CurrencyPanel()
        {
            InitializeComponent();
            this._apiClient = new HttpClient { BaseAddress = new Uri("https://api.frankfurter.app/") };
            Configure();
        }

        #region Initialization
        private void Configure()
        {
            _dataCurrencies.Columns.Clear();
            _dataCurrencies.Columns.Add("Key", "Currency");
            _dataCurrencies.Columns.Add("Value", "Equivalence");
            _dateTimeFrom.Value = DateTime.Today.AddDays(-30);
            _dateTimeTo.Value = DateTime.Today;
            _ = LoadSelectors();
            LoadColorPicker(comboHeaderFill);
            LoadColorPicker(comboBodyFill);
            LoadColorPicker(comboHeaderBorderColor);
            LoadColorPicker(comboBodyBorderColor);
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

        private void LoadColorPicker(System.Windows.Forms.ComboBox combo)
        {
            var colors = typeof(Color).GetProperties()
                            .Where(x => x.PropertyType == typeof(Color))
                            .Select(x => x.GetValue(null)).ToList();
            colors.Insert(0, Color.Empty);
            combo.DataSource = colors;
            combo.MaxDropDownItems = 10;
            combo.IntegralHeight = false;
            combo.DrawMode = DrawMode.OwnerDrawFixed;
            combo.DropDownStyle = ComboBoxStyle.DropDownList;
            combo.DrawItem += (s, e) =>
            {
                e.DrawBackground();
                if (e.Index >= 0)
                {
                    var txt = combo.GetItemText(combo.Items[e.Index]);
                    var color = (Color)combo.Items[e.Index];
                    var r1 = new Rectangle(e.Bounds.Left + 1, e.Bounds.Top + 1,
                        2 * (e.Bounds.Height - 2), e.Bounds.Height - 2);
                    var r2 = Rectangle.FromLTRB(r1.Right + 2, e.Bounds.Top,
                        e.Bounds.Right, e.Bounds.Bottom);
                    if (!color.IsEmpty)
                    {
                        using (var b = new SolidBrush(color))
                            e.Graphics.FillRectangle(b, r1);
                        e.Graphics.DrawRectangle(Pens.Black, r1);
                    }
                    TextRenderer.DrawText(e.Graphics, txt, comboHeaderFill.Font, r2,
                        combo.ForeColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
                }
            };
        }

        private void OnDrawComboHeaderFill(object sender, DrawItemEventArgs e)
        {
            e.DrawBackground();
            if (e.Index >= 0)
            {
                var txt = comboHeaderFill.GetItemText(comboHeaderFill.Items[e.Index]);
                var color = (Color)comboHeaderFill.Items[e.Index];
                var r1 = new Rectangle(e.Bounds.Left + 1, e.Bounds.Top + 1,
                    2 * (e.Bounds.Height - 2), e.Bounds.Height - 2);
                var r2 = Rectangle.FromLTRB(r1.Right + 2, e.Bounds.Top,
                    e.Bounds.Right, e.Bounds.Bottom);
                using (var b = new SolidBrush(color))
                    e.Graphics.FillRectangle(b, r1);
                e.Graphics.DrawRectangle(Pens.Black, r1);
                TextRenderer.DrawText(e.Graphics, txt, comboHeaderFill.Font, r2,
                    comboHeaderFill.ForeColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
            }
        }


        #endregion

        #region Building Sheet
        private void BuildExcel(List<ExchangeRecord> records)
        {
            List<string> targetCurrencies = GetSelectedTargetCurrencies();

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
            table.Options.DateTimeFormat = CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern;
            table.BodyStyle.Border.SetBorderType(CellStyling.BorderStyle.BorderType.Thin);
            table.BodyStyle.NumberingPattern = "0.000";
            ApplyHeaderStyle(table);
            ApplyBodyStyle(table);


            string fileName = GetPathSave();
            if (fileName != null)
            {
                table.Create(fileName);
                MessageBox.Show("File created");
            }
        }
        private void ApplyHeaderStyle(TabularSheet<ExchangeRecord> table)
        {
            Style header = table.HeaderStyle;
            Style header = table.HeaderStyle;
            if ((Color)comboHeaderFill.SelectedItem != Color.Empty)
                header.Fill.BackgroundColor = (Color)comboHeaderFill.SelectedItem;
            if (_applyFontHeader)
            {
                header.Font.FontName = dialogFontHeader.Font.Name;
                header.Font.Bold = dialogFontHeader.Font.Bold;
                header.Font.Italic = dialogFontHeader.Font.Italic;
                header.Font.Color = dialogFontHeader.Color;
            }
        }

        private void ApplyBodyStyle(TabularSheet<ExchangeRecord> table)
        {
            Style body = table.BodyStyle;
            if ((Color)comboBodyFill.SelectedItem != Color.Empty)
                body.Fill.BackgroundColor = (Color)comboBodyFill.SelectedItem;

            if (_applyFontBody)
            {
                body.Font.FontName = dialogFontBody.Font.Name;
                body.Font.Bold = dialogFontBody.Font.Bold;
                body.Font.Italic = dialogFontBody.Font.Italic;
                body.Font.Color = dialogFontBody.Color;
            }
        }
        #endregion

        #region UIControls

        #region Currency selection and visualization

        private List<string> GetSelectedTargetCurrencies()
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

        private async void OnClickSearch(object sender, EventArgs e)
        {
            Currency selected = _comboOriginCurrency.SelectedItem as Currency;
            string currencyCode = selected.CurrencyCode;
            List<string> to = GetSelectedTargetCurrencies();
            var latest = await GetExchangeLatest(currencyCode, to.ToArray());
            FillGrid(latest);
        }

        private void OnClickSelectAll(object sender, EventArgs e)
        {
            for (int i = 0; i < _listDestinationCurrencies.Items.Count; i++)
            {
                _listDestinationCurrencies.SetItemChecked(i, true);
            }
        }

        private void OnClickUnselectAll(object sender, EventArgs e)
        {
            for (int i = 0; i < _listDestinationCurrencies.Items.Count; i++)
            {
                _listDestinationCurrencies.SetItemChecked(i, false);
            }
        }
        #endregion

        #region Export
        private async void OnClickExportSheetHistory(object sender, EventArgs e)
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

            List<string> targetCurrencies = GetSelectedTargetCurrencies();
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

        #endregion

        #region Configuration Styles
        private void OnClickConfigureHeaderFont(object sender, EventArgs e)
        {
            var result = dialogFontHeader.ShowDialog();
            if (result == DialogResult.OK)
                _applyFontHeader = true;
            else
                _applyFontHeader = false;
        }

        private void OnClickConfigureBodyFont(object sender, EventArgs e)
        {
            var result = dialogFontBody.ShowDialog();
            if (result == DialogResult.OK)
                _applyFontBody = true;
            else
                _applyFontBody = false;
        }
        #endregion

        #endregion

        #region API Requests
        private async Task<List<Currency>> GetCurrencies()
        {
            var response = await _apiClient.GetAsync("currencies");
            string currenciesString = await response.Content.ReadAsStringAsync();
            var currencies = JsonConvert.DeserializeObject<Dictionary<string, string>>(currenciesString);
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
            var contentStr = await response.Content.ReadAsStringAsync();
            var content = JsonConvert.DeserializeObject<CurrencyExchangeLatestDTO>(contentStr);
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
            var historicalStr = await response.Content.ReadAsStringAsync();

            var historical = JsonConvert.DeserializeObject<CurrencyExchangeHistoricalDTO>(historicalStr);
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
    }
}
