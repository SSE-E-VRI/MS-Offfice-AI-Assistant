using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using MistralOfficeAddin.Core;
using MistralOfficeAddin.Providers;
using Newtonsoft.Json;

namespace MistralOfficeAddin.UI
{
    public partial class SettingsWindow : Window
    {
        private bool _isKeyRevealed = false;
        private CancellationTokenSource _testCts;
        private CancellationTokenSource _discoverCts;
        private AIProviderType _selectedProviderType = AIProviderType.Mistral;
        private bool _isLoading = false;

        // In-memory working copies per provider during dialog session
        private readonly Dictionary<AIProviderType, ProviderSettings> _workingSettings = new Dictionary<AIProviderType, ProviderSettings>();

        public SettingsWindow()
        {
            InitializeComponent();
            LoadCurrentSettings();
        }

        private void LoadCurrentSettings()
        {
            _isLoading = true;
            try
            {
                var config = ConfigManager.Instance;

                // Clone settings into working copies
                _workingSettings[AIProviderType.Mistral] = CloneSettings(config.Mistral, "https://api.mistral.ai/v1", "mistral-large-latest");
                _workingSettings[AIProviderType.Groq] = CloneSettings(config.Groq, "https://api.groq.com/openai/v1", "llama-3.3-70b-versatile");
                _workingSettings[AIProviderType.Gemini] = CloneSettings(config.Gemini, "https://generativelanguage.googleapis.com", "gemini-1.5-flash");
                _workingSettings[AIProviderType.Custom] = CloneSettings(config.Custom, "http://localhost:11434/v1", "llama3");

                _selectedProviderType = config.ActiveProvider;

                // Select matching provider in ComboBox
                foreach (ComboBoxItem item in CmbProvider.Items)
                {
                    string tag = item.Tag as string;
                    if (string.Equals(tag, _selectedProviderType.ToString(), StringComparison.OrdinalIgnoreCase))
                    {
                        CmbProvider.SelectedItem = item;
                        break;
                    }
                }

                SldTemp.Value = config.Temperature;
                LblTemp.Text = config.Temperature.ToString("0.00");

                SldMaxTokens.Value = config.MaxTokens;
                LblMaxTokens.Text = config.MaxTokens.ToString();

                TxtSystemPrompt.Text = config.SystemPrompt;

                PopulateProviderFields(_selectedProviderType);
            }
            finally
            {
                _isLoading = false;
            }
        }

        private ProviderSettings CloneSettings(ProviderSettings src, string defaultUrl, string defaultModel)
        {
            if (src == null) return new ProviderSettings(defaultUrl, defaultModel);
            var cloned = new ProviderSettings(
                !string.IsNullOrWhiteSpace(src.BaseUrl) ? src.BaseUrl : defaultUrl,
                !string.IsNullOrWhiteSpace(src.DefaultModel) ? src.DefaultModel : defaultModel,
                src.ApiKey);

            if (src.CustomHeaders != null)
            {
                cloned.CustomHeaders = new Dictionary<string, string>(src.CustomHeaders, StringComparer.OrdinalIgnoreCase);
            }
            return cloned;
        }

        private void SaveCurrentFieldsToWorkingCopy()
        {
            if (_isLoading) return;

            ProviderSettings current;
            if (_workingSettings.TryGetValue(_selectedProviderType, out current))
            {
                current.ApiKey = GetCurrentApiKey();
                current.BaseUrl = TxtBaseUrl.Text.Trim();
                current.DefaultModel = CmbModel.Text.Trim();

                if (_selectedProviderType == AIProviderType.Custom && !string.IsNullOrWhiteSpace(TxtCustomHeaders.Text))
                {
                    try
                    {
                        var headers = JsonConvert.DeserializeObject<Dictionary<string, string>>(TxtCustomHeaders.Text);
                        if (headers != null) current.CustomHeaders = headers;
                    }
                    catch { }
                }
            }
        }

        private void PopulateProviderFields(AIProviderType providerType)
        {
            ProviderSettings settings;
            if (!_workingSettings.TryGetValue(providerType, out settings)) return;

            TxtApiKey.Password = settings.ApiKey ?? string.Empty;
            TxtApiKeyPlain.Text = settings.ApiKey ?? string.Empty;
            TxtBaseUrl.Text = settings.BaseUrl ?? string.Empty;

            PnlCustomHeaders.Visibility = (providerType == AIProviderType.Custom) ? Visibility.Visible : Visibility.Collapsed;
            if (providerType == AIProviderType.Custom && settings.CustomHeaders != null && settings.CustomHeaders.Count > 0)
            {
                TxtCustomHeaders.Text = JsonConvert.SerializeObject(settings.CustomHeaders, Formatting.Indented);
            }
            else
            {
                TxtCustomHeaders.Text = string.Empty;
            }

            // Populate Model list options
            CmbModel.Items.Clear();
            List<AIModelInfo> defaultModels = GetDefaultModels(providerType);
            foreach (var m in defaultModels)
            {
                CmbModel.Items.Add(m.Id);
            }

            string modelToSelect = !string.IsNullOrWhiteSpace(settings.DefaultModel) ? settings.DefaultModel : (defaultModels.Count > 0 ? defaultModels[0].Id : "");
            CmbModel.Text = modelToSelect;

            // URL Read-only hints
            if (providerType == AIProviderType.Groq || providerType == AIProviderType.Gemini)
            {
                TxtBaseUrl.IsReadOnly = true;
                TxtBaseUrl.Background = new SolidColorBrush(Color.FromRgb(241, 245, 249));
            }
            else
            {
                TxtBaseUrl.IsReadOnly = false;
                TxtBaseUrl.Background = new SolidColorBrush(Colors.White);
            }
        }

        private List<AIModelInfo> GetDefaultModels(AIProviderType providerType)
        {
            switch (providerType)
            {
                case AIProviderType.Groq:
                    return new List<AIModelInfo>
                    {
                        new AIModelInfo("llama-3.3-70b-versatile"),
                        new AIModelInfo("llama-3.1-8b-instant"),
                        new AIModelInfo("llama-3.2-11b-vision-preview"),
                        new AIModelInfo("llama-3.2-90b-vision-preview"),
                        new AIModelInfo("mixtral-8x7b-32768")
                    };
                case AIProviderType.Gemini:
                    return new List<AIModelInfo>
                    {
                        new AIModelInfo("gemini-2.5-flash"),
                        new AIModelInfo("gemini-2.5-pro"),
                        new AIModelInfo("gemini-1.5-flash"),
                        new AIModelInfo("gemini-1.5-pro")
                    };
                case AIProviderType.Custom:
                    return new List<AIModelInfo>
                    {
                        new AIModelInfo("llama3"),
                        new AIModelInfo("mistral"),
                        new AIModelInfo("qwen2.5")
                    };
                case AIProviderType.Mistral:
                default:
                    return new List<AIModelInfo>
                    {
                        new AIModelInfo("mistral-large-latest"),
                        new AIModelInfo("mistral-small-latest"),
                        new AIModelInfo("open-mistral-nemo"),
                        new AIModelInfo("codestral-latest"),
                        new AIModelInfo("pixtral-large-latest")
                    };
            }
        }

        private void CmbProvider_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isLoading || CmbProvider.SelectedItem == null) return;

            SaveCurrentFieldsToWorkingCopy();

            var selectedItem = CmbProvider.SelectedItem as ComboBoxItem;
            if (selectedItem != null && selectedItem.Tag is string)
            {
                AIProviderType parsed;
                if (Enum.TryParse(selectedItem.Tag as string, true, out parsed))
                {
                    _selectedProviderType = parsed;
                    PopulateProviderFields(_selectedProviderType);
                    LblStatus.Text = string.Empty;
                }
            }
        }

        private string GetCurrentApiKey()
        {
            return _isKeyRevealed ? TxtApiKeyPlain.Text.Trim() : TxtApiKey.Password.Trim();
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

        private async void BtnDiscoverModels_Click(object sender, RoutedEventArgs e)
        {
            SaveCurrentFieldsToWorkingCopy();
            string apiKey = GetCurrentApiKey();
            string baseUrl = TxtBaseUrl.Text.Trim();

            BtnDiscoverModels.IsEnabled = false;
            LblStatus.Foreground = new SolidColorBrush(Color.FromRgb(37, 99, 235));
            LblStatus.Text = "Discovering models from provider...";

            if (_discoverCts != null) _discoverCts.Cancel();
            _discoverCts = new CancellationTokenSource();

            try
            {
                using (var provider = ProviderFactory.CreateForTesting(_selectedProviderType, apiKey, baseUrl))
                {
                    var models = await provider.ListModelsAsync(_discoverCts.Token);
                    if (models != null && models.Count > 0)
                    {
                        string currentText = CmbModel.Text;
                        CmbModel.Items.Clear();
                        foreach (var m in models)
                        {
                            CmbModel.Items.Add(m.Id);
                        }
                        if (!string.IsNullOrWhiteSpace(currentText))
                        {
                            CmbModel.Text = currentText;
                        }
                        else if (CmbModel.Items.Count > 0)
                        {
                            CmbModel.SelectedIndex = 0;
                        }
                        LblStatus.Foreground = new SolidColorBrush(Color.FromRgb(5, 150, 105));
                        LblStatus.Text = string.Format("✓ Discovered {0} models successfully.", models.Count);
                    }
                    else
                    {
                        LblStatus.Foreground = new SolidColorBrush(Color.FromRgb(217, 119, 6));
                        LblStatus.Text = "No models returned by discovery endpoint. You may enter model name manually.";
                    }
                }
            }
            catch (Exception ex)
            {
                LblStatus.Foreground = new SolidColorBrush(Color.FromRgb(220, 38, 38));
                LblStatus.Text = string.Format("Discovery failed: {0}", ex.Message);
            }
            finally
            {
                BtnDiscoverModels.IsEnabled = true;
            }
        }

        private async void BtnTest_Click(object sender, RoutedEventArgs e)
        {
            SaveCurrentFieldsToWorkingCopy();
            string apiKey = GetCurrentApiKey();
            string baseUrl = TxtBaseUrl.Text.Trim();
            string model = CmbModel.Text.Trim();

            if (string.IsNullOrWhiteSpace(apiKey) && _selectedProviderType != AIProviderType.Custom)
            {
                LblStatus.Foreground = new SolidColorBrush(Color.FromRgb(220, 38, 38));
                LblStatus.Text = "Please enter an API key first.";
                return;
            }

            BtnTest.IsEnabled = false;
            LblStatus.Foreground = new SolidColorBrush(Color.FromRgb(37, 99, 235));
            LblStatus.Text = string.Format("Testing connection to {0}...", _selectedProviderType);

            if (_testCts != null) _testCts.Cancel();
            _testCts = new CancellationTokenSource();

            try
            {
                using (var provider = ProviderFactory.CreateForTesting(_selectedProviderType, apiKey, baseUrl, model))
                {
                    bool ok = await provider.TestConnectionAsync(_testCts.Token);
                    if (ok)
                    {
                        LblStatus.Foreground = new SolidColorBrush(Color.FromRgb(5, 150, 105));
                        LblStatus.Text = string.Format("✓ Connection successful! {0} is ready.", _selectedProviderType);
                    }
                }
            }
            catch (Exception ex)
            {
                LblStatus.Foreground = new SolidColorBrush(Color.FromRgb(220, 38, 38));
                LblStatus.Text = string.Format("✗ Connection failed: {0}", ex.Message);
            }
            finally
            {
                BtnTest.IsEnabled = true;
            }
        }

        private void SldTemp_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (LblTemp != null) LblTemp.Text = e.NewValue.ToString("0.00");
        }

        private void SldMaxTokens_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (LblMaxTokens != null) LblMaxTokens.Text = ((int)e.NewValue).ToString();
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            SaveCurrentFieldsToWorkingCopy();

            var config = ConfigManager.Instance;
            config.ActiveProvider = _selectedProviderType;

            ProviderSettings mistralWorking;
            if (_workingSettings.TryGetValue(AIProviderType.Mistral, out mistralWorking))
                config.Mistral = mistralWorking;

            ProviderSettings groqWorking;
            if (_workingSettings.TryGetValue(AIProviderType.Groq, out groqWorking))
                config.Groq = groqWorking;

            ProviderSettings geminiWorking;
            if (_workingSettings.TryGetValue(AIProviderType.Gemini, out geminiWorking))
                config.Gemini = geminiWorking;

            ProviderSettings customWorking;
            if (_workingSettings.TryGetValue(AIProviderType.Custom, out customWorking))
                config.Custom = customWorking;

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
