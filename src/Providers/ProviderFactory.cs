using System;
using MistralOfficeAddin.Core;

namespace MistralOfficeAddin.Providers
{
    public static class ProviderFactory
    {
        public static IAIProvider CreateFromConfig(ConfigManager config)
        {
            if (config == null) config = ConfigManager.Instance;

            switch (config.ActiveProvider)
            {
                case AIProviderType.Groq:
                    var groqCfg = config.Groq ?? new ProviderSettings("https://api.groq.com/openai/v1", "openai/gpt-oss-20b");
                    return new GroqProvider(groqCfg.ApiKey, groqCfg.BaseUrl, groqCfg.DefaultModel);

                case AIProviderType.Gemini:
                    var geminiCfg = config.Gemini ?? new ProviderSettings("https://generativelanguage.googleapis.com", "gemini-3.6-flash");
                    return new GeminiProvider(geminiCfg.ApiKey, geminiCfg.BaseUrl, geminiCfg.DefaultModel);

                case AIProviderType.Custom:
                    var customCfg = config.Custom ?? new ProviderSettings("http://localhost:11434/v1", "llama3");
                    return new CustomApiProvider(
                        customCfg.BaseUrl,
                        customCfg.ApiKey,
                        customCfg.DefaultModel,
                        customCfg.CustomHeaders);

                case AIProviderType.Mistral:
                default:
                    var mistralCfg = config.Mistral ?? new ProviderSettings("https://api.mistral.ai/v1", "mistral-large-latest");
                    return new MistralProvider(mistralCfg.BaseUrl, mistralCfg.ApiKey, mistralCfg.DefaultModel);
            }
        }

        public static IAIProvider CreateForTesting(
            AIProviderType providerType,
            string apiKey,
            string baseUrl = null,
            string defaultModel = null,
            System.Collections.Generic.Dictionary<string, string> customHeaders = null)
        {
            switch (providerType)
            {
                case AIProviderType.Groq:
                    return new GroqProvider(apiKey, baseUrl, defaultModel);

                case AIProviderType.Gemini:
                    return new GeminiProvider(apiKey, baseUrl, defaultModel);

                case AIProviderType.Custom:
                    return new CustomApiProvider(baseUrl, apiKey, defaultModel, customHeaders);

                case AIProviderType.Mistral:
                default:
                    return new MistralProvider(baseUrl, apiKey, defaultModel);
            }
        }
    }
}
