using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using MistralOfficeAddin.API;
using MistralOfficeAddin.API.Models;
using MistralOfficeAddin.Core;
using Newtonsoft.Json;

namespace MistralOfficeAddin.Providers
{
    public class MistralProvider : IAIProvider
    {
        private readonly string _baseUrl;
        private readonly string _apiKey;
        private readonly MistralClient _client;
        private bool _disposed;

        public AIProviderType ProviderType
        {
            get { return AIProviderType.Mistral; }
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

        public MistralProvider(string baseUrl, string apiKey)
        {
            _baseUrl = !string.IsNullOrWhiteSpace(baseUrl) ? baseUrl.TrimEnd('/') : "https://api.mistral.ai/v1";
            _apiKey = apiKey ?? string.Empty;
            _client = new MistralClient(_baseUrl, _apiKey);
        }

        public async Task<bool> TestConnectionAsync(CancellationToken ct = default(CancellationToken))
        {
            try
            {
                return await _client.TestConnectionAsync(ct).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                Logger.Error("MistralProvider: TestConnectionAsync failed", ex);
                throw new AIException(ex.Message, AIProviderType.Mistral, 0, ex);
            }
        }

        public async Task<List<AIModelInfo>> ListModelsAsync(CancellationToken ct = default(CancellationToken))
        {
            var results = new List<AIModelInfo>();
            try
            {
                using (var http = new HttpClient { Timeout = TimeSpan.FromSeconds(15) })
                {
                    if (!string.IsNullOrWhiteSpace(_apiKey))
                    {
                        http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);
                    }
                    string url = string.Format("{0}/models", _baseUrl);
                    using (var resp = await http.GetAsync(url, ct).ConfigureAwait(false))
                    {
                        if (resp.IsSuccessStatusCode)
                        {
                            string json = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                            var modelResp = JsonConvert.DeserializeObject<MistralModelsListResponse>(json);
                            if (modelResp != null && modelResp.Data != null)
                            {
                                foreach (var m in modelResp.Data)
                                {
                                    if (!string.IsNullOrWhiteSpace(m.Id))
                                    {
                                        bool isVision = CheckVisionSupport(m.Id);
                                        results.Add(new AIModelInfo(m.Id, m.Id, isVision));
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("MistralProvider: ListModelsAsync discovery failed: {0}", ex.Message));
            }

            if (results.Count == 0)
            {
                results.Add(new AIModelInfo("mistral-large-latest", "mistral-large-latest", false));
                results.Add(new AIModelInfo("mistral-small-latest", "mistral-small-latest", false));
                results.Add(new AIModelInfo("open-mistral-nemo", "open-mistral-nemo", false));
                results.Add(new AIModelInfo("codestral-latest", "codestral-latest", false));
                results.Add(new AIModelInfo("pixtral-large-latest", "pixtral-large-latest (Vision)", true));
                results.Add(new AIModelInfo("pixtral-12b-2409", "pixtral-12b-2409 (Vision)", true));
            }

            return results;
        }

        public async Task<AIResponse> ChatAsync(AIRequest request, CancellationToken ct = default(CancellationToken))
        {
            if (request == null) throw new ArgumentNullException("request");
            try
            {
                string text = await _client.ChatAsync(
                    request.Model,
                    request.Messages,
                    request.Temperature,
                    request.MaxTokens,
                    ct).ConfigureAwait(false);

                return new AIResponse(text) { Model = request.Model };
            }
            catch (Exception ex)
            {
                Logger.Error("MistralProvider: ChatAsync error", ex);
                throw new AIException(ex.Message, AIProviderType.Mistral, 0, ex);
            }
        }

        public async Task StreamChatAsync(AIRequest request, Action<string> onDeltaReceived, CancellationToken ct = default(CancellationToken))
        {
            if (request == null) throw new ArgumentNullException("request");
            if (onDeltaReceived == null) throw new ArgumentNullException("onDeltaReceived");

            try
            {
                await _client.StreamChatCallbackAsync(
                    request.Model,
                    request.Messages,
                    request.Temperature,
                    request.MaxTokens,
                    onDeltaReceived,
                    ct).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                Logger.Error("MistralProvider: StreamChatAsync error", ex);
                throw new AIException(ex.Message, AIProviderType.Mistral, 0, ex);
            }
        }

        public bool CheckVisionSupport(string model)
        {
            if (string.IsNullOrWhiteSpace(model)) return false;
            return model.IndexOf("pixtral", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   model.IndexOf("vision", StringComparison.OrdinalIgnoreCase) >= 0;
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

        private class MistralModelsListResponse
        {
            [JsonProperty("data")]
            public List<MistralModelItem> Data { get; set; }
        }

        private class MistralModelItem
        {
            [JsonProperty("id")]
            public string Id { get; set; }
        }
    }
}
