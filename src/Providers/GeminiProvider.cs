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
    public class GeminiProvider : IAIProvider
    {
        private const string DefaultBaseUrl = "https://generativelanguage.googleapis.com";
        private readonly string _apiKey;
        private readonly string _baseUrl;
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

        public GeminiProvider(string apiKey, string baseUrl = null)
        {
            _apiKey = apiKey ?? string.Empty;
            _baseUrl = !string.IsNullOrWhiteSpace(baseUrl) ? baseUrl.TrimEnd('/') : DefaultBaseUrl;

            var handler = new HttpClientHandler
            {
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
            };

            _httpClient = new HttpClient(handler)
            {
                Timeout = TimeSpan.FromSeconds(120)
            };
            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        }

        public async Task<bool> TestConnectionAsync(CancellationToken ct = default(CancellationToken))
        {
            try
            {
                string url = string.Format("{0}/v1beta/models?key={1}", _baseUrl, Uri.EscapeDataString(_apiKey));
                using (var resp = await _httpClient.GetAsync(url, ct).ConfigureAwait(false))
                {
                    if (resp.IsSuccessStatusCode)
                    {
                        return true;
                    }

                    string errBody = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                    Logger.Warn(string.Format("GeminiProvider: TestConnection failed HTTP {0}: {1}", (int)resp.StatusCode, errBody));
                    throw new AIException(
                        string.Format("HTTP {0} ({1}): {2}", (int)resp.StatusCode, resp.ReasonPhrase, errBody),
                        AIProviderType.Gemini,
                        (int)resp.StatusCode);
                }
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
                                    if (canGenerate)
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
                list.Add(new AIModelInfo("gemini-2.5-flash", "gemini-2.5-flash", true));
                list.Add(new AIModelInfo("gemini-2.5-pro", "gemini-2.5-pro", true));
                list.Add(new AIModelInfo("gemini-1.5-flash", "gemini-1.5-flash", true));
                list.Add(new AIModelInfo("gemini-1.5-pro", "gemini-1.5-pro", true));
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
            string json = JsonConvert.SerializeObject(payload);

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
                throw new AIException(string.Format("Gemini API returned HTTP {0} ({1}). Check log for details.", code, resp.ReasonPhrase), AIProviderType.Gemini, code);
            }
        }

        public async Task StreamChatAsync(AIRequest request, Action<string> onDeltaReceived, CancellationToken ct = default(CancellationToken))
        {
            if (request == null) throw new ArgumentNullException("request");
            if (onDeltaReceived == null) throw new ArgumentNullException("onDeltaReceived");

            string modelId = CleanModelId(request.Model);
            string url = string.Format("{0}/v1beta/models/{1}:streamGenerateContent?alt=sse&key={2}",
                _baseUrl, Uri.EscapeDataString(modelId), Uri.EscapeDataString(_apiKey));

            var payload = BuildGeminiPayload(request);
            string json = JsonConvert.SerializeObject(payload);

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
                    throw new AIException(string.Format("Gemini streaming returned HTTP {0} ({1}). Check log for details.", statusCode, response.ReasonPhrase), AIProviderType.Gemini, statusCode);
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
            if (string.IsNullOrWhiteSpace(model)) return "gemini-1.5-flash";
            model = model.Trim();
            if (model.StartsWith("models/")) return model.Substring(7);
            return model;
        }

        private object BuildGeminiPayload(AIRequest request)
        {
            var contents = new List<object>();

            foreach (var m in request.Messages)
            {
                if (m.IsSystem) continue; // Passed via systemInstruction

                string role = m.IsUser ? "user" : "model";
                var parts = new List<object>();

                if (!string.IsNullOrEmpty(m.Content))
                {
                    parts.Add(new { text = m.Content });
                }

                contents.Add(new
                {
                    role = role,
                    parts = parts
                });
            }

            // Append active image attachments to last user content block if present
            if (request.Attachments != null && request.Attachments.Count > 0)
            {
                var lastUserBlock = contents.FindLast(c => ((dynamic)c).role == "user");
                if (lastUserBlock != null)
                {
                    var partsList = (List<object>)((dynamic)lastUserBlock).parts;
                    foreach (var att in request.Attachments)
                    {
                        if (att.IsImage && att.RawBytes != null && att.RawBytes.Length > 0)
                        {
                            partsList.Add(new
                            {
                                inlineData = new
                                {
                                    mimeType = !string.IsNullOrWhiteSpace(att.ContentType) ? att.ContentType : "image/jpeg",
                                    data = Convert.ToBase64String(att.RawBytes)
                                }
                            });
                        }
                    }
                }
            }

            var generationConfig = new
            {
                temperature = request.Temperature,
                maxOutputTokens = request.MaxTokens > 0 ? request.MaxTokens : 4096
            };

            if (!string.IsNullOrWhiteSpace(request.SystemPrompt))
            {
                var systemInstruction = new
                {
                    parts = new object[]
                    {
                        new { text = request.SystemPrompt }
                    }
                };

                return new
                {
                    contents = contents,
                    generationConfig = generationConfig,
                    systemInstruction = systemInstruction
                };
            }

            return new
            {
                contents = contents,
                generationConfig = generationConfig
            };
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
