using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using MSOfficeAIAssistant.Core;

namespace MSOfficeAIAssistant.Providers
{
    public class GroqProvider : IAIProvider
    {
        private const string DefaultBaseUrl = "https://api.groq.com/openai/v1";
        private readonly string _apiKey;
        private readonly string _testModel;
        private readonly OpenAICompatibleClient _client;
        private bool _disposed;

        public AIProviderType ProviderType
        {
            get { return AIProviderType.Groq; }
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

        public GroqProvider(string apiKey, string baseUrl = null, string testModel = null)
        {
            _apiKey = apiKey ?? string.Empty;
            _testModel = !string.IsNullOrWhiteSpace(testModel) ? testModel.Trim() : "openai/gpt-oss-20b";
            string endpoint = !string.IsNullOrWhiteSpace(baseUrl) ? baseUrl.TrimEnd('/') : DefaultBaseUrl;
            _client = new OpenAICompatibleClient(endpoint, _apiKey, AIProviderType.Groq);
        }

        public async Task<bool> TestConnectionAsync(CancellationToken ct = default(CancellationToken))
        {
            return await _client.TestConnectionAsync(_testModel, ct).ConfigureAwait(false);
        }

        public async Task<List<AIModelInfo>> ListModelsAsync(CancellationToken ct = default(CancellationToken))
        {
            var list = await _client.ListModelsAsync(ct).ConfigureAwait(false);
            if (list != null && list.Count > 0)
            {
                var filtered = new List<AIModelInfo>();
                foreach (var m in list)
                {
                    if (m.Id.IndexOf("whisper", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        m.Id.IndexOf("tts", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        m.Id.IndexOf("embed", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        continue;
                    }
                    filtered.Add(m);
                }
                if (filtered.Count > 0) return filtered;
            }

            return new List<AIModelInfo>
            {
                new AIModelInfo("openai/gpt-oss-20b", "openai/gpt-oss-20b (Fast)", false),
                new AIModelInfo("openai/gpt-oss-120b", "openai/gpt-oss-120b (Recommended)", false),
                new AIModelInfo("qwen/qwen3.6-27b", "qwen/qwen3.6-27b", false)
            };
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
                   model.IndexOf("llava", StringComparison.OrdinalIgnoreCase) >= 0;
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
