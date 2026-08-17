using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using MistralOfficeAddin.Core;

namespace MistralOfficeAddin.Providers
{
    public class GroqProvider : IAIProvider
    {
        private const string GroqBaseUrl = "https://api.groq.com/openai/v1";
        private readonly string _apiKey;
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
                       AICapabilities.ConnectionTest;
            }
        }

        public GroqProvider(string apiKey)
        {
            _apiKey = apiKey ?? string.Empty;
            _client = new OpenAICompatibleClient(GroqBaseUrl, _apiKey, AIProviderType.Groq);
        }

        public async Task<bool> TestConnectionAsync(CancellationToken ct = default(CancellationToken))
        {
            return await _client.TestConnectionAsync("llama-3.1-8b-instant", ct).ConfigureAwait(false);
        }

        public async Task<List<AIModelInfo>> ListModelsAsync(CancellationToken ct = default(CancellationToken))
        {
            var list = await _client.ListModelsAsync(ct).ConfigureAwait(false);
            if (list == null || list.Count == 0)
            {
                list = new List<AIModelInfo>
                {
                    new AIModelInfo("llama-3.3-70b-versatile", "llama-3.3-70b-versatile (Recommended)", false),
                    new AIModelInfo("llama-3.1-8b-instant", "llama-3.1-8b-instant (Fast)", false),
                    new AIModelInfo("llama-3.2-11b-vision-preview", "llama-3.2-11b-vision-preview (Vision)", true),
                    new AIModelInfo("llama-3.2-90b-vision-preview", "llama-3.2-90b-vision-preview (Vision)", true),
                    new AIModelInfo("mixtral-8x7b-32768", "mixtral-8x7b-32768", false),
                    new AIModelInfo("gemma2-9b-it", "gemma2-9b-it", false)
                };
            }
            return list;
        }

        public async Task<AIResponse> ChatAsync(AIRequest request, CancellationToken ct = default(CancellationToken))
        {
            if (request == null) throw new ArgumentNullException("request");
            string content = await _client.ChatAsync(
                request.Model,
                request.Messages,
                request.Temperature,
                request.MaxTokens,
                ct).ConfigureAwait(false);

            return new AIResponse(content) { Model = request.Model };
        }

        public async Task StreamChatAsync(AIRequest request, Action<string> onDeltaReceived, CancellationToken ct = default(CancellationToken))
        {
            if (request == null) throw new ArgumentNullException("request");
            await _client.StreamChatCallbackAsync(
                request.Model,
                request.Messages,
                request.Temperature,
                request.MaxTokens,
                onDeltaReceived,
                ct).ConfigureAwait(false);
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
