using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Newtonsoft.Json;

namespace MistralOfficeAddin.Core
{
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

        public string BaseUrl { get; set; }
        public string ApiKey { get; set; }
        public string DefaultModel { get; set; }
        public double Temperature { get; set; }
        public int MaxTokens { get; set; }
        public string SystemPrompt { get; set; }
        public bool AutoInsertResponse { get; set; }

        public ConfigManager()
        {
            BaseUrl = "https://api.mistral.ai/v1";
            ApiKey = string.Empty;
            DefaultModel = "mistral-large-latest";
            Temperature = 0.7;
            MaxTokens = 4096;
            SystemPrompt = "You are an expert AI assistant embedded inside Microsoft Office. Help the user draft, refine, summarize, and analyze documents with professional quality.";
            AutoInsertResponse = false;

            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            _configDirectory = Path.Combine(localAppData, "MistralOfficeAddin");
            _configFilePath = Path.Combine(_configDirectory, "config.dat");
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
                        BaseUrl = this.BaseUrl,
                        ApiKey = this.ApiKey,
                        DefaultModel = this.DefaultModel,
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
                        this.BaseUrl = !string.IsNullOrWhiteSpace(dto.BaseUrl) ? dto.BaseUrl : "https://api.mistral.ai/v1";
                        this.ApiKey = dto.ApiKey ?? string.Empty;
                        this.DefaultModel = !string.IsNullOrWhiteSpace(dto.DefaultModel) ? dto.DefaultModel : "mistral-large-latest";
                        this.Temperature = dto.Temperature;
                        this.MaxTokens = dto.MaxTokens > 0 ? dto.MaxTokens : 4096;
                        this.SystemPrompt = !string.IsNullOrWhiteSpace(dto.SystemPrompt) ? dto.SystemPrompt : this.SystemPrompt;
                        this.AutoInsertResponse = dto.AutoInsertResponse;
                        Logger.Info("Configuration successfully loaded and decrypted.");
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error("Failed to load configuration", ex);
                }
            }
        }

        private class ConfigDto
        {
            public string BaseUrl { get; set; }
            public string ApiKey { get; set; }
            public string DefaultModel { get; set; }
            public double Temperature { get; set; }
            public int MaxTokens { get; set; }
            public string SystemPrompt { get; set; }
            public bool AutoInsertResponse { get; set; }
        }
    }
}
