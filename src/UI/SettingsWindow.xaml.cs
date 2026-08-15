using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using MistralOfficeAddin.API;
using MistralOfficeAddin.Core;

namespace MistralOfficeAddin.UI
{
    public partial class SettingsWindow : Window
    {
        private bool _isKeyRevealed = false;
        private CancellationTokenSource _testCts;

        public SettingsWindow()
        {
            InitializeComponent();
            LoadCurrentSettings();
        }

        private void LoadCurrentSettings()
        {
            var config = ConfigManager.Instance;
            TxtApiKey.Password = config.ApiKey;
            TxtApiKeyPlain.Text = config.ApiKey;
            TxtBaseUrl.Text = config.BaseUrl;

            // Select Model
            bool modelMatched = false;
            foreach (ComboBoxItem item in CmbModel.Items)
            {
                if (string.Equals(Convert.ToString(item.Content), config.DefaultModel, StringComparison.OrdinalIgnoreCase))
                {
                    CmbModel.SelectedItem = item;
                    modelMatched = true;
                    break;
                }
            }
            if (!modelMatched && !string.IsNullOrEmpty(config.DefaultModel))
            {
                var customItem = new ComboBoxItem { Content = config.DefaultModel };
                CmbModel.Items.Add(customItem);
                CmbModel.SelectedItem = customItem;
            }

            SldTemp.Value = config.Temperature;
            LblTemp.Text = config.Temperature.ToString("0.00");

            SldMaxTokens.Value = config.MaxTokens;
            LblMaxTokens.Text = config.MaxTokens.ToString();

            TxtSystemPrompt.Text = config.SystemPrompt;
        }

        private void SldTemp_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (LblTemp != null)
            {
                LblTemp.Text = e.NewValue.ToString("0.00");
            }
        }

        private void SldMaxTokens_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (LblMaxTokens != null)
            {
                LblMaxTokens.Text = ((int)e.NewValue).ToString();
            }
        }

        private void BtnToggleKey_Click(object sender, RoutedEventArgs e)
        {
            _isKeyRevealed = !_isKeyRevealed;
            if (_isKeyRevealed)
            {
                TxtApiKeyPlain.Text = TxtApiKey.Password;
                TxtApiKey.Visibility = Visibility.Collapsed;
                TxtApiKeyPlain.Visibility = Visibility.Visible;
                BtnToggleKey.Content = "🔒";
            }
            else
            {
                TxtApiKey.Password = TxtApiKeyPlain.Text;
                TxtApiKeyPlain.Visibility = Visibility.Collapsed;
                TxtApiKey.Visibility = Visibility.Visible;
                BtnToggleKey.Content = "👁";
            }
        }

        private string GetCurrentApiKey()
        {
            return _isKeyRevealed ? TxtApiKeyPlain.Text.Trim() : TxtApiKey.Password.Trim();
        }

        private async void BtnTest_Click(object sender, RoutedEventArgs e)
        {
            string key = GetCurrentApiKey();
            string baseUrl = TxtBaseUrl.Text.Trim();

            if (string.IsNullOrWhiteSpace(key))
            {
                LblStatus.Foreground = new SolidColorBrush(Color.FromRgb(220, 38, 38)); // Red
                LblStatus.Text = "Please enter an API key first.";
                return;
            }

            BtnTest.IsEnabled = false;
            LblStatus.Foreground = new SolidColorBrush(Color.FromRgb(37, 99, 235)); // Blue
            LblStatus.Text = "Testing connection to Mistral AI...";

            if (_testCts != null)
            {
                _testCts.Cancel();
            }
            _testCts = new CancellationTokenSource();

            try
            {
                using (var client = new MistralClient(baseUrl, key))
                {
                    bool ok = await client.TestConnectionAsync(_testCts.Token);
                    if (ok)
                    {
                        LblStatus.Foreground = new SolidColorBrush(Color.FromRgb(5, 150, 105)); // Green
                        LblStatus.Text = "✓ Connection successful! API key is valid.";
                    }
                }
            }
            catch (Exception ex)
            {
                LblStatus.Foreground = new SolidColorBrush(Color.FromRgb(220, 38, 38)); // Red
                LblStatus.Text = string.Format("✗ Connection failed: {0}", ex.Message);
            }
            finally
            {
                BtnTest.IsEnabled = true;
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            var config = ConfigManager.Instance;
            config.ApiKey = GetCurrentApiKey();
            config.BaseUrl = TxtBaseUrl.Text.Trim();

            if (CmbModel.SelectedItem is ComboBoxItem)
            {
                var selectedItem = (ComboBoxItem)CmbModel.SelectedItem;
                config.DefaultModel = Convert.ToString(selectedItem.Content);
            }

            config.Temperature = Math.Round(SldTemp.Value, 2);
            config.MaxTokens = (int)SldMaxTokens.Value;
            config.SystemPrompt = TxtSystemPrompt.Text.Trim();

            config.Save();
            this.DialogResult = true;
            this.Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}
