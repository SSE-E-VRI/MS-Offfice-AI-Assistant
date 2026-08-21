using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using MSOfficeAIAssistant.Core;

namespace MSOfficeAIAssistant.Providers
{
    public class CustomApiProvider : IAIProvider
    {
        private readonly string _baseUrl;
        private readonly string _apiKey;
        private readonly string _defaultModel;
        private readonly OpenAICompatibleClient _client;
        private bool _disposed;

        public AIProviderType ProviderType
        {
            get { return AIProviderType.Custom; }
        }

        public AICapabilities Capabilities
        {
            get
            {
                return AICapabilities.Chat |
                       AICapabilities.Streaming |
                       AICapabilities.Vision |
                       AICapabilities.ModelListing |
                       AICapabilities.ConnectionTest |
                       AICapabilities.StructuredOutput |
                       AICapabilities.ToolCalling |
                       AICapabilities.JsonMode;
            }
        }

        public CustomApiProvider(string baseUrl, string apiKey, string defaultModel = null, Dictionary<string, string> customHeaders = null)
        {
            _baseUrl = !string.IsNullOrWhiteSpace(baseUrl) ? baseUrl.TrimEnd('/') : "http://localhost:11434/v1";
            _apiKey = apiKey ?? string.Empty;
            _defaultModel = !string.IsNullOrWhiteSpace(defaultModel) ? defaultModel : "llama3";

            // Enforce HTTPS for non-loopback endpoints
            _client = new OpenAICompatibleClient(_baseUrl, _apiKey, AIProviderType.Custom, customHeaders, enforceHttps: true);
        }

        public async Task<bool> TestConnectionAsync(CancellationToken ct = default(CancellationToken))
        {
            return await _client.TestConnectionAsync(_defaultModel, ct).ConfigureAwait(false);
        }

        public async Task<List<AIModelInfo>> ListModelsAsync(CancellationToken ct = default(CancellationToken))
        {
            var list = await _client.ListModelsAsync(ct).ConfigureAwait(false);
            if (list == null || list.Count == 0)
            {
                list = new List<AIModelInfo>
                {
                    new AIModelInfo(_defaultModel, _defaultModel, CheckVisionSupport(_defaultModel))
                };
            }
            return list;
        }

        public async Task<AIResponse> ChatAsync(AIRequest request, CancellationToken ct = default(CancellationToken))
        {
            if (request == null) throw new ArgumentNullException("request");
            string content = await _client.ChatAsync(request, ct).ConfigureAwait(false);

            return new AIResponse(content) { Model = request.Model };
        }

        public async Task StreamChatAsync(AIRequest request, Action<string> onDeltaReceived, CancellationToken ct = default(CancellationToken))
        {
            if (request == null) throw new ArgumentNullException("request");
            await _client.StreamChatCallbackAsync(request, onDeltaReceived, ct).ConfigureAwait(false);
        }

        public bool CheckVisionSupport(string model)
        {
            if (string.IsNullOrWhiteSpace(model)) return false;
            return model.IndexOf("vision", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   model.IndexOf("llava", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   model.IndexOf("pixtral", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                if (_client != null)
                {
                    _client.Dispose();
                }
                _disposed = true;
            }
        }
    }
}
