using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using MistralOfficeAddin.API.Models;
using MistralOfficeAddin.Core;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace MistralOfficeAddin.Providers
{
    public class GeminiPart
    {
        [JsonProperty("text", NullValueHandling = NullValueHandling.Ignore)]
        public string Text { get; set; }

        [JsonProperty("inlineData", NullValueHandling = NullValueHandling.Ignore)]
        public GeminiInlineData InlineData { get; set; }
    }

    public class GeminiInlineData
    {
        [JsonProperty("mimeType")]
        public string MimeType { get; set; }

        [JsonProperty("data")]
        public string Data { get; set; }
    }

    public class GeminiContent
    {
        [JsonProperty("role")]
        public string Role { get; set; }

        [JsonProperty("parts")]
        public List<GeminiPart> Parts { get; set; }

        public GeminiContent()
        {
            Parts = new List<GeminiPart>();
        }
    }

    public class GeminiGenerationConfig
    {
        [JsonProperty("maxOutputTokens")]
        public int MaxOutputTokens { get; set; }
    }

    public class GeminiSystemInstruction
    {
        [JsonProperty("parts")]
        public List<GeminiPart> Parts { get; set; }

        public GeminiSystemInstruction()
        {
            Parts = new List<GeminiPart>();
        }
    }

    public class GeminiRequestPayload
    {
        [JsonProperty("contents")]
        public List<GeminiContent> Contents { get; set; }

        [JsonProperty("generationConfig", NullValueHandling = NullValueHandling.Ignore)]
        public GeminiGenerationConfig GenerationConfig { get; set; }

        [JsonProperty("systemInstruction", NullValueHandling = NullValueHandling.Ignore)]
        public GeminiSystemInstruction SystemInstruction { get; set; }
    }

    public class GeminiProvider : IAIProvider
    {
        private const string DefaultBaseUrl = "https://generativelanguage.googleapis.com";
        private readonly string _apiKey;
        private readonly string _baseUrl;
        private readonly string _testModel;
        private readonly HttpClient _httpClient;
        private bool _disposed;

        public AIProviderType ProviderType
        {
            get { return AIProviderType.Gemini; }
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

        public GeminiProvider(string apiKey, string baseUrl = null, string testModel = null)
        {
            _apiKey = apiKey ?? string.Empty;
            _baseUrl = !string.IsNullOrWhiteSpace(baseUrl) ? baseUrl.TrimEnd('/') : DefaultBaseUrl;
            _testModel = CleanModelId(testModel);

            var handler = new HttpClientHandler
            {
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
            };

            _httpClient = new HttpClient(handler)
            {
                Timeout = TimeSpan.FromSeconds(120)
            };
            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            if (!string.IsNullOrWhiteSpace(_apiKey))
            {
                _httpClient.DefaultRequestHeaders.TryAddWithoutValidation("x-goog-api-key", _apiKey);
            }
        }

        public async Task<bool> TestConnectionAsync(CancellationToken ct = default(CancellationToken))
        {
            try
            {
                return await CanGenerateWithModelAsync(_testModel, ct).ConfigureAwait(false);
            }
            catch (AIException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Logger.Error("GeminiProvider: TestConnection exception", ex);
                throw new AIException(ex.Message, AIProviderType.Gemini, 0, ex);
            }
        }

        /// <summary>
        /// Returns only models this API key can actually call. Model listing alone does not
        /// guarantee free-tier or project-level generation access.
        /// </summary>
        public async Task<List<AIModelInfo>> FindWorkingModelsAsync(CancellationToken ct = default(CancellationToken))
        {
            var available = await ListModelsAsync(ct).ConfigureAwait(false);
            var working = new List<AIModelInfo>();

            foreach (var model in available)
            {
                ct.ThrowIfCancellationRequested();
                try
                {
                    if (await CanGenerateWithModelAsync(model.Id, ct).ConfigureAwait(false))
                    {
                        working.Add(model);
                    }
                }
                catch (AIException)
                {
                    // Try the next model; access and free-tier availability are key-specific.
                }
            }

            return working;
        }

        public async Task<List<AIModelInfo>> ListModelsAsync(CancellationToken ct = default(CancellationToken))
        {
            var list = new List<AIModelInfo>();
            try
            {
                string url = string.Format("{0}/v1beta/models?key={1}", _baseUrl, Uri.EscapeDataString(_apiKey));
                using (var resp = await _httpClient.GetAsync(url, ct).ConfigureAwait(false))
                {
                    if (resp.IsSuccessStatusCode)
                    {
                        string json = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                        var jObj = JObject.Parse(json);
                        var modelsArr = jObj["models"] as JArray;
                        if (modelsArr != null)
                        {
                            foreach (var m in modelsArr)
                            {
                                string name = (string)m["name"]; // e.g. "models/gemini-1.5-flash"
                                if (!string.IsNullOrEmpty(name))
                                {
                                    string modelId = name.StartsWith("models/") ? name.Substring(7) : name;
                                    string dispName = (string)m["displayName"] ?? modelId;
                                    var supportedMethods = m["supportedGenerationMethods"] as JArray;
                                    bool canGenerate = false;
                                    if (supportedMethods != null)
                                    {
                                        foreach (var method in supportedMethods)
                                        {
                                            if ((string)method == "generateContent") canGenerate = true;
                                        }
                                    }
                                    if (canGenerate && modelId.StartsWith("gemini", StringComparison.OrdinalIgnoreCase))
                                    {
                                        list.Add(new AIModelInfo(modelId, dispName, CheckVisionSupport(modelId)));
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("GeminiProvider: ListModelsAsync failed: {0}", ex.Message));
            }

            if (list.Count == 0)
            {
                list.Add(new AIModelInfo("gemini-3.6-flash", "gemini-3.6-flash (Recommended)", true));
            }

            return list;
        }

        public async Task<AIResponse> ChatAsync(AIRequest request, CancellationToken ct = default(CancellationToken))
        {
            if (request == null) throw new ArgumentNullException("request");
            string modelId = CleanModelId(request.Model);

            string url = string.Format("{0}/v1beta/models/{1}:generateContent?key={2}",
                _baseUrl, Uri.EscapeDataString(modelId), Uri.EscapeDataString(_apiKey));

            var payload = BuildGeminiPayload(request);
            string json = JsonConvert.SerializeObject(payload, Formatting.None);

            using (var content = new StringContent(json, Encoding.UTF8, "application/json"))
            using (var resp = await _httpClient.PostAsync(url, content, ct).ConfigureAwait(false))
            {
                if (resp.IsSuccessStatusCode)
                {
                    string respBody = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                    string text = ExtractTextFromGeminiResponse(respBody);
                    return new AIResponse(text) { Model = modelId };
                }

                int code = (int)resp.StatusCode;
                string err = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                Logger.Warn(string.Format("Gemini API error (HTTP {0}): {1}", code, err));
                throw new AIException(string.Format("Gemini API returned HTTP {0}: {1}", code, err), AIProviderType.Gemini, code);
            }
        }

        public async Task StreamChatAsync(AIRequest request, Action<string> onDeltaReceived, CancellationToken ct = default(CancellationToken))
        {
            if (request == null) throw new ArgumentNullException("request");
            if (onDeltaReceived == null) throw new ArgumentNullException("onDeltaReceived");

            // Gemini's SSE connection can remain open without a terminal event on some
            // free-tier models. Use the standard generateContent endpoint so the Office UI
            // always receives a completed response instead of waiting indefinitely.
            var completeResponse = await ChatAsync(request, ct).ConfigureAwait(false);
            if (completeResponse != null && !string.IsNullOrEmpty(completeResponse.Content))
            {
                onDeltaReceived(completeResponse.Content);
            }
        }

        private async Task StreamChatViaSseAsync(AIRequest request, Action<string> onDeltaReceived, CancellationToken ct)
        {
            if (request == null) throw new ArgumentNullException("request");
            if (onDeltaReceived == null) throw new ArgumentNullException("onDeltaReceived");

            string modelId = CleanModelId(request.Model);
            string url = string.Format("{0}/v1beta/models/{1}:streamGenerateContent?alt=sse&key={2}",
                _baseUrl, Uri.EscapeDataString(modelId), Uri.EscapeDataString(_apiKey));

            var payload = BuildGeminiPayload(request);
            string json = JsonConvert.SerializeObject(payload, Formatting.None);

            var reqMessage = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new StringContent(json, Encoding.UTF8, "application/json")
            };

            using (var response = await _httpClient.SendAsync(reqMessage, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false))
            {
                if (!response.IsSuccessStatusCode)
                {
                    int statusCode = (int)response.StatusCode;
                    string errorBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    Logger.Warn(string.Format("Gemini streaming error (HTTP {0}): {1}", statusCode, errorBody));
                    throw new AIException(string.Format("Gemini streaming returned HTTP {0}: {1}", statusCode, errorBody), AIProviderType.Gemini, statusCode);
                }

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
                        if (GeminiStreamingParser.TryParseLine(line, out delta, out isDone))
                        {
                            if (!string.IsNullOrEmpty(delta))
                            {
                                onDeltaReceived(delta);
                            }
                            if (isDone) break;
                        }
                    }
                }
            }
        }

        public bool CheckVisionSupport(string model)
        {
            // All modern Gemini 1.5 and 2.x models natively support vision/multimodal input
            return true;
        }

        private static string CleanModelId(string model)
        {
            if (string.IsNullOrWhiteSpace(model)) return "gemini-3.6-flash";
            model = model.Trim();
            if (model.StartsWith("models/")) return model.Substring(7);
            return model;
        }

        private async Task<bool> CanGenerateWithModelAsync(string model, CancellationToken ct)
        {
            string modelId = CleanModelId(model);
            string url = string.Format("{0}/v1beta/models/{1}:generateContent?key={2}",
                _baseUrl, Uri.EscapeDataString(modelId), Uri.EscapeDataString(_apiKey));

            var payload = new GeminiRequestPayload
            {
                Contents = new List<GeminiContent>
                {
                    new GeminiContent
                    {
                        Role = "user",
                        Parts = new List<GeminiPart> { new GeminiPart { Text = "Reply with OK." } }
                    }
                },
                GenerationConfig = new GeminiGenerationConfig { MaxOutputTokens = 16 }
            };

            string json = JsonConvert.SerializeObject(payload, Formatting.None);
            using (var content = new StringContent(json, Encoding.UTF8, "application/json"))
            using (var resp = await _httpClient.PostAsync(url, content, ct).ConfigureAwait(false))
            {
                if (resp.IsSuccessStatusCode) return true;

                string error = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                Logger.Warn(string.Format("GeminiProvider: Model {0} cannot generate (HTTP {1}): {2}", modelId, (int)resp.StatusCode, error));
                throw new AIException(
                    string.Format("Gemini model {0} returned HTTP {1}: {2}", modelId, (int)resp.StatusCode, error),
                    AIProviderType.Gemini,
                    (int)resp.StatusCode);
            }
        }

        private GeminiRequestPayload BuildGeminiPayload(AIRequest request)
        {
            var contents = new List<GeminiContent>();
            string systemPrompt = request.SystemPrompt;

            if (request.Messages != null)
            {
                foreach (var m in request.Messages)
                {
                    if (m.IsSystem)
                    {
                        if (string.IsNullOrWhiteSpace(systemPrompt))
                        {
                            systemPrompt = m.Content;
                        }
                        continue;
                    }

                    string role = m.IsUser ? "user" : "model";
                    var parts = new List<GeminiPart>();

                    if (!string.IsNullOrEmpty(m.Content))
                    {
                        parts.Add(new GeminiPart { Text = m.Content });
                    }

                    contents.Add(new GeminiContent
                    {
                        Role = role,
                        Parts = parts
                    });
                }
            }

            // Append active image attachments to last user content block
            if (request.Attachments != null && request.Attachments.Count > 0)
            {
                var lastUserBlock = contents.FindLast(c => c.Role == "user");
                if (lastUserBlock == null)
                {
                    lastUserBlock = new GeminiContent { Role = "user" };
                    contents.Add(lastUserBlock);
                }

                foreach (var att in request.Attachments)
                {
                    if (att.IsImage && att.RawBytes != null && att.RawBytes.Length > 0)
                    {
                        lastUserBlock.Parts.Add(new GeminiPart
                        {
                            InlineData = new GeminiInlineData
                            {
                                MimeType = !string.IsNullOrWhiteSpace(att.ContentType) ? att.ContentType : "image/jpeg",
                                Data = Convert.ToBase64String(att.RawBytes)
                            }
                        });
                    }
                }
            }

            // Ensure every content block has at least one part (Gemini API requires non-empty parts)
            foreach (var c in contents)
            {
                if (c.Parts.Count == 0)
                {
                    c.Parts.Add(new GeminiPart { Text = " " });
                }
            }

            // If contents is completely empty, add a default user message
            if (contents.Count == 0)
            {
                contents.Add(new GeminiContent
                {
                    Role = "user",
                    Parts = new List<GeminiPart> { new GeminiPart { Text = "Hello" } }
                });
            }

            var payload = new GeminiRequestPayload
            {
                Contents = contents,
                GenerationConfig = new GeminiGenerationConfig
                {
                    MaxOutputTokens = request.MaxTokens > 0 ? request.MaxTokens : 4096
                }
            };

            if (!string.IsNullOrWhiteSpace(systemPrompt))
            {
                payload.SystemInstruction = new GeminiSystemInstruction
                {
                    Parts = new List<GeminiPart>
                    {
                        new GeminiPart { Text = systemPrompt }
                    }
                };
            }

            return payload;
        }

        private static string ExtractTextFromGeminiResponse(string json)
        {
            try
            {
                var jObj = JObject.Parse(json);
                var candidates = jObj["candidates"] as JArray;
                if (candidates != null && candidates.Count > 0)
                {
                    var content = candidates[0]["content"];
                    if (content != null)
                    {
                        var parts = content["parts"] as JArray;
                        if (parts != null && parts.Count > 0)
                        {
                            var sb = new StringBuilder();
                            foreach (var p in parts)
                            {
                                sb.Append((string)p["text"]);
                            }
                            return sb.ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("GeminiProvider: Failed to parse response: {0}", ex.Message));
            }
            return string.Empty;
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
    }
}
