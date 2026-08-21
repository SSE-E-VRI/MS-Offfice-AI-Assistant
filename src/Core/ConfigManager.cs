using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using MSOfficeAIAssistant.Providers;
using Newtonsoft.Json;

namespace MSOfficeAIAssistant.Core
{
    public class ProviderSettings
    {
        public string ApiKey { get; set; }
        public string BaseUrl { get; set; }
        public string DefaultModel { get; set; }
        public Dictionary<string, string> CustomHeaders { get; set; }

        public ProviderSettings()
        {
            ApiKey = string.Empty;
            BaseUrl = string.Empty;
            DefaultModel = string.Empty;
            CustomHeaders = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        public ProviderSettings(string baseUrl, string defaultModel, string apiKey = "")
        {
            ApiKey = apiKey ?? string.Empty;
            BaseUrl = baseUrl ?? string.Empty;
            DefaultModel = defaultModel ?? string.Empty;
            CustomHeaders = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }
    }

    public class ConfigManager
    {
        private static readonly object _lock = new object();
        private static ConfigManager _instance;
        private readonly string _configDirectory;
        private readonly string _configFilePath;

        public static ConfigManager Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                        {
                            _instance = new ConfigManager();
                            _instance.Load();
                        }
                    }
                }
                return _instance;
            }
        }

        public AIProviderType ActiveProvider { get; set; }
        public ProviderSettings Mistral { get; set; }
        public ProviderSettings Groq { get; set; }
        public ProviderSettings Gemini { get; set; }
        public ProviderSettings Custom { get; set; }

        public double Temperature { get; set; }
        public int MaxTokens { get; set; }
        public string SystemPrompt { get; set; }
        public bool AutoInsertResponse { get; set; }
        public bool LoadFailed { get; private set; }

        // Backward-compatibility accessors
        public string BaseUrl
        {
            get { return GetActiveProviderSettings().BaseUrl; }
            set { GetActiveProviderSettings().BaseUrl = value; }
        }

        public string ApiKey
        {
            get { return GetActiveProviderSettings().ApiKey; }
            set { GetActiveProviderSettings().ApiKey = value; }
        }

        public string DefaultModel
        {
            get { return GetActiveProviderSettings().DefaultModel; }
            set { GetActiveProviderSettings().DefaultModel = value; }
        }

        public ConfigManager()
        {
            ActiveProvider = AIProviderType.Mistral;

            Mistral = new ProviderSettings("https://api.mistral.ai/v1", "mistral-large-latest");
            Groq = new ProviderSettings("https://api.groq.com/openai/v1", "openai/gpt-oss-20b");
            Gemini = new ProviderSettings("https://generativelanguage.googleapis.com", "gemini-3.6-flash");
            Custom = new ProviderSettings("http://localhost:11434/v1", "llama3");

            Temperature = 0.7;
            MaxTokens = 4096;
            SystemPrompt = "You are an expert AI assistant embedded inside Microsoft Office. Help the user draft, refine, summarize, and analyze documents with professional quality.";
            AutoInsertResponse = false;

            _configDirectory = AppPaths.DataDirectory;
            _configFilePath = Path.Combine(_configDirectory, "config.dat");
        }

        public ProviderSettings GetActiveProviderSettings()
        {
            switch (ActiveProvider)
            {
                case AIProviderType.Groq:
                    return Groq ?? (Groq = new ProviderSettings("https://api.groq.com/openai/v1", "openai/gpt-oss-20b"));
                case AIProviderType.Gemini:
                    return Gemini ?? (Gemini = new ProviderSettings("https://generativelanguage.googleapis.com", "gemini-3.6-flash"));
                case AIProviderType.Custom:
                    return Custom ?? (Custom = new ProviderSettings("http://localhost:11434/v1", "llama3"));
                case AIProviderType.Mistral:
                default:
                    return Mistral ?? (Mistral = new ProviderSettings("https://api.mistral.ai/v1", "mistral-large-latest"));
            }
        }

        public ProviderSettings GetSettingsForProvider(AIProviderType providerType)
        {
            switch (providerType)
            {
                case AIProviderType.Groq: return Groq;
                case AIProviderType.Gemini: return Gemini;
                case AIProviderType.Custom: return Custom;
                case AIProviderType.Mistral:
                default: return Mistral;
            }
        }

        public void Save()
        {
            lock (_lock)
            {
                try
                {
                    if (!Directory.Exists(_configDirectory))
                    {
                        Directory.CreateDirectory(_configDirectory);
                    }

                    var dto = new ConfigDto
                    {
                        ActiveProvider = this.ActiveProvider.ToString(),
                        Mistral = this.Mistral,
                        Groq = this.Groq,
                        Gemini = this.Gemini,
                        Custom = this.Custom,
                        Temperature = this.Temperature,
                        MaxTokens = this.MaxTokens,
                        SystemPrompt = this.SystemPrompt,
                        AutoInsertResponse = this.AutoInsertResponse
                    };

                    string json = JsonConvert.SerializeObject(dto, Formatting.Indented);
                    byte[] plainBytes = Encoding.UTF8.GetBytes(json);
                    byte[] encryptedBytes = ProtectedData.Protect(plainBytes, null, DataProtectionScope.CurrentUser);

                    File.WriteAllBytes(_configFilePath, encryptedBytes);
                    Logger.Info("Configuration securely saved to config.dat using DPAPI.");
                }
                catch (Exception ex)
                {
                    Logger.Error("Failed to save configuration", ex);
                }
            }
        }

        public void Load()
        {
            lock (_lock)
            {
                try
                {
                    if (!File.Exists(_configFilePath))
                    {
                        Logger.Info("No existing config file found. Using default settings.");
                        return;
                    }

                    byte[] encryptedBytes = File.ReadAllBytes(_configFilePath);
                    if (encryptedBytes.Length == 0) return;

                    byte[] plainBytes = ProtectedData.Unprotect(encryptedBytes, null, DataProtectionScope.CurrentUser);
                    string json = Encoding.UTF8.GetString(plainBytes);

                    var dto = JsonConvert.DeserializeObject<ConfigDto>(json);
                    if (dto != null)
                    {
                        // Migration: Check if legacy flat configuration is present
                        if (dto.Mistral == null && (!string.IsNullOrWhiteSpace(dto.LegacyApiKey) || !string.IsNullOrWhiteSpace(dto.LegacyBaseUrl)))
                        {
                            Logger.Info("Migrating legacy single-provider config into Mistral provider slot.");
                            this.Mistral = new ProviderSettings(
                                !string.IsNullOrWhiteSpace(dto.LegacyBaseUrl) ? dto.LegacyBaseUrl : "https://api.mistral.ai/v1",
                                !string.IsNullOrWhiteSpace(dto.LegacyDefaultModel) ? dto.LegacyDefaultModel : "mistral-large-latest",
                                dto.LegacyApiKey ?? string.Empty);
                            this.ActiveProvider = AIProviderType.Mistral;
                        }
                        else
                        {
                            if (!string.IsNullOrWhiteSpace(dto.ActiveProvider))
                            {
                                AIProviderType parsedProvider;
                                if (Enum.TryParse(dto.ActiveProvider, true, out parsedProvider))
                                {
                                    this.ActiveProvider = parsedProvider;
                                }
                            }

                            if (dto.Mistral != null) this.Mistral = dto.Mistral;
                            if (dto.Groq != null) this.Groq = dto.Groq;
                            if (dto.Gemini != null) this.Gemini = dto.Gemini;
                            if (dto.Custom != null) this.Custom = dto.Custom;
                        }

                        if (this.Mistral == null) this.Mistral = new ProviderSettings("https://api.mistral.ai/v1", "mistral-large-latest");
                        if (this.Groq == null) this.Groq = new ProviderSettings("https://api.groq.com/openai/v1", "openai/gpt-oss-20b");
                        if (this.Gemini == null) this.Gemini = new ProviderSettings("https://generativelanguage.googleapis.com", "gemini-3.6-flash");
                        MigrateRetiredProviderModels();
                        if (this.Custom == null) this.Custom = new ProviderSettings("http://localhost:11434/v1", "llama3");

                        this.Temperature = dto.Temperature.HasValue ? dto.Temperature.Value : 0.7;
                        this.MaxTokens = dto.MaxTokens > 0 ? dto.MaxTokens : 4096;
                        this.SystemPrompt = !string.IsNullOrWhiteSpace(dto.SystemPrompt) ? dto.SystemPrompt : this.SystemPrompt;
                        this.AutoInsertResponse = dto.AutoInsertResponse;

                        Logger.Info("Configuration successfully loaded and decrypted.");
                    }
                }
                catch (Exception ex)
                {
                    this.LoadFailed = true;
                    Logger.Error("Failed to load configuration", ex);
                }
            }
        }

        private void MigrateRetiredProviderModels()
        {
            if (Groq != null && IsOneOf(Groq.DefaultModel,
                "llama-3.3-70b-versatile", "llama-3.1-8b-instant", "llama-3.2-11b-vision-preview",
                "llama-3.2-90b-vision-preview", "qwen/qwen3-32b", "deepseek-r1-distill-llama-70b",
                "mixtral-8x7b-32768"))
            {
                Groq.DefaultModel = "openai/gpt-oss-20b";
                Logger.Info("Migrated retired Groq default model to openai/gpt-oss-20b.");
            }

            if (Gemini != null && IsOneOf(Gemini.DefaultModel,
                "gemini-1.5-flash", "gemini-1.5-pro", "gemini-2.0-flash", "gemini-2.0-flash-lite",
                "gemini-2.5-flash", "gemini-2.5-flash-lite", "gemini-2.5-pro"))
            {
                Gemini.DefaultModel = "gemini-3.6-flash";
                Logger.Info("Migrated retired Gemini default model to gemini-3.6-flash.");
            }
        }

        private static bool IsOneOf(string value, params string[] candidates)
        {
            if (string.IsNullOrWhiteSpace(value)) return false;
            foreach (string candidate in candidates)
            {
                if (string.Equals(value, candidate, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }

        private class ConfigDto
        {
            public string ActiveProvider { get; set; }
            public ProviderSettings Mistral { get; set; }
            public ProviderSettings Groq { get; set; }
            public ProviderSettings Gemini { get; set; }
            public ProviderSettings Custom { get; set; }

            // Legacy fields for backward compatibility migration
            [JsonProperty("BaseUrl")]
            public string LegacyBaseUrl { get; set; }

            [JsonProperty("ApiKey")]
            public string LegacyApiKey { get; set; }

            [JsonProperty("DefaultModel")]
            public string LegacyDefaultModel { get; set; }

            public double? Temperature { get; set; }
            public int MaxTokens { get; set; }
            public string SystemPrompt { get; set; }
            public bool AutoInsertResponse { get; set; }
        }
    }
}
