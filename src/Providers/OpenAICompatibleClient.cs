using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using MSOfficeAIAssistant.API;
using MSOfficeAIAssistant.API.Models;
using MSOfficeAIAssistant.Attachments;
using MSOfficeAIAssistant.Core;
using Newtonsoft.Json;

namespace MSOfficeAIAssistant.Providers
{
    public class OpenAICompatibleClient : IDisposable
    {
        private readonly HttpClient _httpClient;
        private readonly string _baseUrl;
        private readonly string _apiKey;
        private readonly AIProviderType _providerType;
        private bool _disposed;

        public OpenAICompatibleClient(
            string baseUrl,
            string apiKey,
            AIProviderType providerType,
            Dictionary<string, string> customHeaders = null,
            bool enforceHttps = false)
        {
            _providerType = providerType;
            _baseUrl = !string.IsNullOrWhiteSpace(baseUrl) ? baseUrl.TrimEnd('/') : string.Empty;
            _apiKey = apiKey ?? string.Empty;

            if (enforceHttps && !string.IsNullOrWhiteSpace(_baseUrl))
            {
                ValidateEndpointUrl(_baseUrl);
            }

            var handler = new HttpClientHandler
            {
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
            };

            _httpClient = new HttpClient(handler)
            {
                Timeout = TimeSpan.FromSeconds(120)
            };

            if (!string.IsNullOrWhiteSpace(_apiKey))
            {
                _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);
            }

            if (customHeaders != null)
            {
                foreach (var kvp in customHeaders)
                {
                    if (!string.IsNullOrWhiteSpace(kvp.Key) && kvp.Value != null)
                    {
                        _httpClient.DefaultRequestHeaders.TryAddWithoutValidation(kvp.Key, kvp.Value);
                    }
                }
            }

            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        }

        public static void ValidateEndpointUrl(string url)
        {
            Uri uri;
            if (!Uri.TryCreate(url, UriKind.Absolute, out uri))
            {
                throw new ArgumentException("Invalid URL format: " + url);
            }

            if (uri.Scheme.Equals("https", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            // If HTTP, allow only local loopback endpoints (e.g. Ollama, LM Studio, LocalAI)
            string host = uri.Host.ToLowerInvariant();
            bool isLoopback = uri.IsLoopback ||
                              host == "localhost" ||
                              host == "127.0.0.1" ||
                              host == "::1" ||
                              host == "0.0.0.0";

            if (!isLoopback)
            {
                throw new ArgumentException("Remote AI endpoints require HTTPS for security. Non-HTTPS is only permitted for local loopback (localhost/127.0.0.1).");
            }
        }

        public async Task<bool> TestConnectionAsync(string model = null, CancellationToken ct = default(CancellationToken))
        {
            try
            {
                string testModel = !string.IsNullOrWhiteSpace(model) ? model : "gpt-3.5-turbo";
                var testMessages = new List<ChatMessage> { new ChatMessage("user", "hi") };

                var requestPayload = new ChatRequest
                {
                    Model = testModel,
                    Messages = testMessages,
                    MaxTokens = 2,
                    Stream = false
                };

                string url = string.Format("{0}/chat/completions", _baseUrl);
                string json = JsonConvert.SerializeObject(requestPayload);

                using (var content = new StringContent(json, Encoding.UTF8, "application/json"))
                using (var response = await _httpClient.PostAsync(url, content, ct).ConfigureAwait(false))
                {
                    if (response.IsSuccessStatusCode)
                    {
                        return true;
                    }

                    string errBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    Logger.Warn(string.Format("OpenAICompatibleClient TestConnection failed HTTP {0}: {1}", (int)response.StatusCode, errBody));
                    throw new AIException(
                        string.Format("HTTP {0} ({1}): {2}", (int)response.StatusCode, response.ReasonPhrase, errBody),
                        _providerType,
                        (int)response.StatusCode);
                }
            }
            catch (AIException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Logger.Error("OpenAICompatibleClient TestConnection error", ex);
                throw new AIException(ex.Message, _providerType, 0, ex);
            }
        }

        public async Task<List<AIModelInfo>> ListModelsAsync(CancellationToken ct = default(CancellationToken))
        {
            var list = new List<AIModelInfo>();
            try
            {
                string url = string.Format("{0}/models", _baseUrl);
                using (var response = await _httpClient.GetAsync(url, ct).ConfigureAwait(false))
                {
                    if (response.IsSuccessStatusCode)
                    {
                        string json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                        var parsed = JsonConvert.DeserializeObject<OpenAIModelListResponse>(json);
                        if (parsed != null && parsed.Data != null)
                        {
                            foreach (var item in parsed.Data)
                            {
                                if (!string.IsNullOrWhiteSpace(item.Id))
                                {
                                    bool isVision = item.Id.IndexOf("vision", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                                    item.Id.IndexOf("llava", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                                    item.Id.IndexOf("pixtral", StringComparison.OrdinalIgnoreCase) >= 0;
                                    list.Add(new AIModelInfo(item.Id, item.Id, isVision));
                                }
                            }
                        }
                    }
                    else
                    {
                        string errBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                        Logger.Warn(string.Format("OpenAICompatibleClient: ListModelsAsync failed HTTP {0}: {1}", (int)response.StatusCode, errBody));
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("OpenAICompatibleClient: ListModelsAsync failed: {0}", ex.Message));
            }
            return list;
        }

        public static object BuildPayload(
            string model,
            List<ChatMessage> messages,
            double temperature,
            int maxTokens,
            bool stream,
            List<AttachmentBlock> attachments = null,
            string systemPrompt = null,
            string responseFormat = null,
            object tools = null,
            object toolChoice = null,
            Dictionary<string, object> extraParameters = null)
        {
            var messagePayloads = new List<object>();

            if (!string.IsNullOrWhiteSpace(systemPrompt))
            {
                bool hasSystemMessage = messages != null && messages.Any(m => m.IsSystem);
                if (!hasSystemMessage)
                {
                    messagePayloads.Add(new { role = "system", content = systemPrompt });
                }
            }

            if (messages != null)
            {
                int lastUserIdx = -1;
                for (int i = messages.Count - 1; i >= 0; i--)
                {
                    if (messages[i].IsUser)
                    {
                        lastUserIdx = i;
                        break;
                    }
                }

                bool hasImages = attachments != null && attachments.Any(a => a.IsImage && a.RawBytes != null && a.RawBytes.Length > 0);

                for (int i = 0; i < messages.Count; i++)
                {
                    var m = messages[i];
                    if (i == lastUserIdx && hasImages)
                    {
                        var contentParts = new List<object>();
                        if (!string.IsNullOrEmpty(m.Content))
                        {
                            contentParts.Add(new { type = "text", text = m.Content });
                        }

                        foreach (var att in attachments)
                        {
                            if (att.IsImage && att.RawBytes != null && att.RawBytes.Length > 0)
                            {
                                string mime = !string.IsNullOrWhiteSpace(att.ContentType) ? att.ContentType : "image/jpeg";
                                string dataUrl = string.Format("data:{0};base64,{1}", mime, Convert.ToBase64String(att.RawBytes));
                                contentParts.Add(new
                                {
                                    type = "image_url",
                                    image_url = new { url = dataUrl }
                                });
                            }
                        }

                        messagePayloads.Add(new
                        {
                            role = m.Role,
                            content = contentParts
                        });
                    }
                    else
                    {
                        messagePayloads.Add(new
                        {
                            role = m.Role,
                            content = m.Content ?? string.Empty
                        });
                    }
                }
            }

            var payload = new Dictionary<string, object>();
            payload["model"] = model;
            payload["messages"] = messagePayloads;
            payload["temperature"] = Math.Max(0.0, Math.Min(2.0, temperature));
            payload["max_tokens"] = maxTokens > 0 ? maxTokens : 4096;
            payload["stream"] = stream;

            if (!string.IsNullOrWhiteSpace(responseFormat))
            {
                if (string.Equals(responseFormat, "json_object", StringComparison.OrdinalIgnoreCase))
                {
                    payload["response_format"] = new Dictionary<string, object> { { "type", "json_object" } };
                }
                else
                {
                    payload["response_format"] = new Dictionary<string, object> { { "type", responseFormat } };
                }
            }

            if (tools != null)
            {
                payload["tools"] = tools;
                if (toolChoice != null)
                {
                    payload["tool_choice"] = toolChoice;
                }
            }

            if (extraParameters != null)
            {
                foreach (var kvp in extraParameters)
                {
                    if (!string.IsNullOrWhiteSpace(kvp.Key) && !payload.ContainsKey(kvp.Key))
                    {
                        payload[kvp.Key] = kvp.Value;
                    }
                }
            }

            return payload;
        }

        public static object BuildPayload(AIRequest request, bool stream)
        {
            if (request == null) return null;
            return BuildPayload(
                request.Model,
                request.Messages,
                request.Temperature,
                request.MaxTokens,
                stream,
                request.Attachments,
                request.SystemPrompt,
                request.ResponseFormat,
                request.Tools,
                request.ToolChoice,
                request.ExtraParameters);
        }

        public async Task<string> ChatAsync(
            string model,
            List<ChatMessage> messages,
            double temperature,
            int maxTokens,
            List<AttachmentBlock> attachments = null,
            CancellationToken ct = default(CancellationToken),
            string systemPrompt = null)
        {
            return await ChatAsync(new AIRequest
            {
                Model = model,
                Messages = messages,
                Temperature = temperature,
                MaxTokens = maxTokens,
                Attachments = attachments,
                SystemPrompt = systemPrompt
            }, ct).ConfigureAwait(false);
        }

        public async Task<string> ChatAsync(AIRequest request, CancellationToken ct = default(CancellationToken))
        {
            if (request == null) throw new ArgumentNullException("request");
            string url = string.Format("{0}/chat/completions", _baseUrl);
            var requestPayload = BuildPayload(request, false);
            string requestJson = JsonConvert.SerializeObject(requestPayload);
            int maxRetries = 3;
            int delayMs = 1000;

            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                ct.ThrowIfCancellationRequested();
                bool shouldRetry = false;

                try
                {
                    using (var content = new StringContent(requestJson, Encoding.UTF8, "application/json"))
                    using (var response = await _httpClient.PostAsync(url, content, ct).ConfigureAwait(false))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            string respBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                            var chatResponse = JsonConvert.DeserializeObject<ChatResponse>(respBody);
                            if (chatResponse != null && chatResponse.Choices != null && chatResponse.Choices.Count > 0)
                            {
                                return chatResponse.Choices[0].Message != null ? (chatResponse.Choices[0].Message.Content ?? string.Empty) : string.Empty;
                            }
                            return string.Empty;
                        }

                        int statusCode = (int)response.StatusCode;
                        string errorBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                        Logger.Warn(string.Format("{0} API error (HTTP {1}): {2}", _providerType, statusCode, errorBody));

                        if ((statusCode == 429 || statusCode >= 500) && attempt < maxRetries)
                        {
                            shouldRetry = true;
                        }
                        else
                        {
                            throw new AIException(
                                string.Format("{0} returned HTTP {1}: {2}", _providerType, statusCode, errorBody),
                                _providerType,
                                statusCode);
                        }
                    }
                }
                catch (HttpRequestException ex)
                {
                    if (attempt < maxRetries) shouldRetry = true;
                    else throw new AIException(ex.Message, _providerType, 0, ex);
                }

                if (shouldRetry)
                {
                    await Task.Delay(delayMs, ct).ConfigureAwait(false);
                    delayMs *= 2;
                }
            }

            throw new AIException("Failed to get response after multiple attempts.", _providerType);
        }

        public async Task StreamChatCallbackAsync(
            string model,
            List<ChatMessage> messages,
            double temperature,
            int maxTokens,
            Action<string> onDeltaReceived,
            List<AttachmentBlock> attachments = null,
            CancellationToken ct = default(CancellationToken),
            string systemPrompt = null)
        {
            await StreamChatCallbackAsync(new AIRequest
            {
                Model = model,
                Messages = messages,
                Temperature = temperature,
                MaxTokens = maxTokens,
                Attachments = attachments,
                SystemPrompt = systemPrompt
            }, onDeltaReceived, ct).ConfigureAwait(false);
        }

        public async Task StreamChatCallbackAsync(
            AIRequest request,
            Action<string> onDeltaReceived,
            CancellationToken ct = default(CancellationToken))
        {
            if (request == null) throw new ArgumentNullException("request");
            if (onDeltaReceived == null) throw new ArgumentNullException("onDeltaReceived");

            string url = string.Format("{0}/chat/completions", _baseUrl);
            var requestPayload = BuildPayload(request, true);
            string requestJson = JsonConvert.SerializeObject(requestPayload);
            int maxRetries = 3;
            int delayMs = 1000;
            bool deltasEmitted = false;

            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                ct.ThrowIfCancellationRequested();
                bool shouldRetry = false;

                try
                {
                    var httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, url)
                    {
                        Content = new StringContent(requestJson, Encoding.UTF8, "application/json")
                    };

                    using (var response = await _httpClient.SendAsync(httpRequestMessage, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false))
                    {
                        if (!response.IsSuccessStatusCode)
                        {
                            int statusCode = (int)response.StatusCode;
                            string errorBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                            Logger.Warn(string.Format("{0} streaming error (HTTP {1}): {2}", _providerType, statusCode, errorBody));

                            if ((statusCode == 429 || statusCode >= 500) && attempt < maxRetries)
                            {
                                shouldRetry = true;
                            }
                            else
                            {
                                throw new AIException(
                                    string.Format("{0} streaming returned HTTP {1}: {2}", _providerType, statusCode, errorBody),
                                    _providerType,
                                    statusCode);
                            }
                        }
                        else
                        {
                            using (var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
                            using (var reader = new StreamReader(stream, Encoding.UTF8))
                            {
                                while (!reader.EndOfStream)
                                {
                                    ct.ThrowIfCancellationRequested();
                                    string line = await reader.ReadLineAsync().ConfigureAwait(false);
                                    if (line == null) break;

                                    string delta;
                                    bool isDone;
                                    if (StreamingParser.TryParseLine(line, out delta, out isDone))
                                    {
                                        if (!string.IsNullOrEmpty(delta))
                                        {
                                            onDeltaReceived(delta);
                                            deltasEmitted = true;
                                        }
                                        if (isDone) break;
                                    }
                                }
                            }
                            return;
                        }
                    }
                }
                catch (HttpRequestException ex)
                {
                    if (deltasEmitted || attempt >= maxRetries)
                        throw new AIException(ex.Message, _providerType, 0, ex);
                    shouldRetry = true;
                }

                if (shouldRetry)
                {
                    await Task.Delay(delayMs, ct).ConfigureAwait(false);
                    delayMs *= 2;
                }
            }

            throw new AIException("Failed to stream response after multiple attempts.", _providerType);
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                if (_httpClient != null)
                {
                    _httpClient.Dispose();
                }
                _disposed = true;
            }
        }

        private class OpenAIModelListResponse
        {
            [JsonProperty("data")]
            public List<OpenAIModelItem> Data { get; set; }
        }

        private class OpenAIModelItem
        {
            [JsonProperty("id")]
            public string Id { get; set; }
        }
    }
}
