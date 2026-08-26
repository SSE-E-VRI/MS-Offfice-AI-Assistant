using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Navigation;
using MSOfficeAIAssistant.Core;
using MSOfficeAIAssistant.Providers;
using Newtonsoft.Json;

namespace MSOfficeAIAssistant.UI
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

        private static void EnsureWpfApplication()
        {
            if (System.Windows.Application.Current != null) return;
            try
            {
                var app = new System.Windows.Application();
                app.ShutdownMode = System.Windows.ShutdownMode.OnExplicitShutdown;
            }
            catch { }
        }

        public SettingsWindow()
        {
            EnsureWpfApplication();
            _isLoading = true;
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
                _workingSettings[AIProviderType.Groq] = CloneSettings(config.Groq, "https://api.groq.com/openai/v1", "openai/gpt-oss-20b");
                _workingSettings[AIProviderType.Gemini] = CloneSettings(config.Gemini, "https://generativelanguage.googleapis.com", "gemini-3.6-flash");
                _workingSettings[AIProviderType.Custom] = CloneSettings(config.Custom, "http://localhost:11434/v1", "llama3");

                _selectedProviderType = config.ActiveProvider;

                if (CmbProvider != null)
                {
                    foreach (ComboBoxItem item in CmbProvider.Items)
                    {
                        string tag = item.Tag as string;
                        if (string.Equals(tag, _selectedProviderType.ToString(), StringComparison.OrdinalIgnoreCase))
                        {
                            CmbProvider.SelectedItem = item;
                            break;
                        }
                    }
                }

                if (SldTemp != null) SldTemp.Value = config.Temperature;
                if (LblTemp != null) LblTemp.Text = config.Temperature.ToString("0.00");

                if (SldMaxTokens != null) SldMaxTokens.Value = config.MaxTokens;
                if (LblMaxTokens != null) LblMaxTokens.Text = config.MaxTokens.ToString();

                if (TxtSystemPrompt != null) TxtSystemPrompt.Text = config.SystemPrompt;

                if (CmbDomainPack != null)
                {
                    foreach (ComboBoxItem item in CmbDomainPack.Items)
                    {
                        string tag = item.Tag as string;
                        if (string.Equals(tag, config.DomainPack, StringComparison.OrdinalIgnoreCase))
                        {
                            CmbDomainPack.SelectedItem = item;
                            break;
                        }
                    }
                    if (CmbDomainPack.SelectedItem == null && CmbDomainPack.Items.Count > 0)
                    {
                        CmbDomainPack.SelectedIndex = 0;
                    }
                }

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
                if (TxtBaseUrl != null) current.BaseUrl = TxtBaseUrl.Text.Trim();
                if (CmbModel != null) current.DefaultModel = CmbModel.Text.Trim();

                if (_selectedProviderType == AIProviderType.Custom && TxtCustomHeaders != null && !string.IsNullOrWhiteSpace(TxtCustomHeaders.Text))
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

            if (TxtApiKey != null) TxtApiKey.Password = settings.ApiKey ?? string.Empty;
            if (TxtApiKeyPlain != null) TxtApiKeyPlain.Text = settings.ApiKey ?? string.Empty;
            if (TxtBaseUrl != null) TxtBaseUrl.Text = settings.BaseUrl ?? string.Empty;

            if (PnlCustomHeaders != null)
            {
                PnlCustomHeaders.Visibility = (providerType == AIProviderType.Custom) ? Visibility.Visible : Visibility.Collapsed;
            }
            if (TxtCustomHeaders != null)
            {
                if (providerType == AIProviderType.Custom && settings.CustomHeaders != null && settings.CustomHeaders.Count > 0)
                {
                    TxtCustomHeaders.Text = JsonConvert.SerializeObject(settings.CustomHeaders, Formatting.Indented);
                }
                else
                {
                    TxtCustomHeaders.Text = string.Empty;
                }
            }

            // Populate Model list options
            if (CmbModel != null)
            {
                CmbModel.Items.Clear();
                List<AIModelInfo> defaultModels = GetDefaultModels(providerType);
                foreach (var m in defaultModels)
                {
                    CmbModel.Items.Add(m.Id);
                }

                string modelToSelect = !string.IsNullOrWhiteSpace(settings.DefaultModel) ? settings.DefaultModel : (defaultModels.Count > 0 ? defaultModels[0].Id : "");
                CmbModel.Text = modelToSelect;
            }

            // URL Read-only hints
            if (TxtBaseUrl != null)
            {
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

            // Update API Key Help Link (contextual sign-up/key registration)
            UpdateApiKeyRegistrationLink(providerType);
        }

        private List<AIModelInfo> GetDefaultModels(AIProviderType providerType)
        {
            switch (providerType)
            {
                case AIProviderType.Groq:
                    return new List<AIModelInfo>
                    {
                        new AIModelInfo("openai/gpt-oss-20b"),
                        new AIModelInfo("openai/gpt-oss-120b"),
                        new AIModelInfo("qwen/qwen3.6-27b")
                    };
                case AIProviderType.Gemini:
                    return new List<AIModelInfo>
                    {
                        new AIModelInfo("gemini-3.6-flash")
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
            if (_isLoading || CmbProvider == null || CmbProvider.SelectedItem == null) return;

            SaveCurrentFieldsToWorkingCopy();

            var selectedItem = CmbProvider.SelectedItem as ComboBoxItem;
            if (selectedItem != null && selectedItem.Tag is string)
            {
                AIProviderType parsed;
                if (Enum.TryParse(selectedItem.Tag as string, true, out parsed))
                {
                    _selectedProviderType = parsed;
                    PopulateProviderFields(_selectedProviderType);
                    if (LblStatus != null)
                    {
                        LblStatus.Text = string.Empty;
                    }
                }
            }
        }

        private string GetCurrentApiKey()
        {
            if (_isKeyRevealed && TxtApiKeyPlain != null) return TxtApiKeyPlain.Text.Trim();
            if (TxtApiKey != null) return TxtApiKey.Password.Trim();
            return string.Empty;
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
            bool verifyGeminiModels = _selectedProviderType == AIProviderType.Gemini;

            BtnDiscoverModels.IsEnabled = false;
            LblStatus.Foreground = new SolidColorBrush(Color.FromRgb(37, 99, 235));
            LblStatus.Text = verifyGeminiModels
                ? "Checking Gemini models available to this API key..."
                : "Discovering models from provider...";

            if (_discoverCts != null) _discoverCts.Cancel();
            _discoverCts = new CancellationTokenSource();

            try
            {
                using (var provider = ProviderFactory.CreateForTesting(_selectedProviderType, apiKey, baseUrl))
                {
                    List<AIModelInfo> models;
                    var geminiProvider = provider as GeminiProvider;
                    if (geminiProvider != null)
                    {
                        models = await geminiProvider.FindWorkingModelsAsync(_discoverCts.Token);
                    }
                    else
                    {
                        models = await provider.ListModelsAsync(_discoverCts.Token);
                    }
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
                        if (verifyGeminiModels)
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

            if (CmbDomainPack != null && CmbDomainPack.SelectedItem != null)
            {
                var selectedItem = CmbDomainPack.SelectedItem as ComboBoxItem;
                if (selectedItem != null && selectedItem.Tag is string)
                {
                    config.DomainPack = selectedItem.Tag as string;
                }
            }

            config.Save();
            this.DialogResult = true;
            this.Close();
        }

        public static string GetProviderRegistrationUrl(AIProviderType providerType)
        {
            switch (providerType)
            {
                case AIProviderType.Mistral:
                    return "https://console.mistral.ai/api-keys/";
                case AIProviderType.Groq:
                    return "https://console.groq.com/keys";
                case AIProviderType.Gemini:
                    return "https://aistudio.google.com/app/apikey";
                case AIProviderType.Custom:
                default:
                    return null;
            }
        }

        private void UpdateApiKeyRegistrationLink(AIProviderType providerType)
        {
            if (PnlApiKeyHelp == null) return;

            string url = GetProviderRegistrationUrl(providerType);
            if (!string.IsNullOrWhiteSpace(url) && providerType != AIProviderType.Custom)
            {
                PnlApiKeyHelp.Visibility = Visibility.Visible;
                if (LnkApiKey != null)
                {
                    try
                    {
                        LnkApiKey.NavigateUri = new Uri(url);
                    }
                    catch
                    {
                        LnkApiKey.NavigateUri = null;
                    }
                }
            }
            else
            {
                PnlApiKeyHelp.Visibility = Visibility.Collapsed;
            }
        }

        private void LnkApiKey_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            try
            {
                string targetUrl = e.Uri != null ? e.Uri.AbsoluteUri : GetProviderRegistrationUrl(_selectedProviderType);
                if (!string.IsNullOrWhiteSpace(targetUrl))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = targetUrl,
                        UseShellExecute = true
                    });
                }
            }
            catch (Exception ex)
            {
                if (LblStatus != null)
                {
                    LblStatus.Foreground = new SolidColorBrush(Color.FromRgb(220, 38, 38));
                    LblStatus.Text = string.Format("Could not open browser: {0}", ex.Message);
                }
            }
            finally
            {
                e.Handled = true;
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}
