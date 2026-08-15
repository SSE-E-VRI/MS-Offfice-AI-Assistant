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

namespace MistralOfficeAddin.API
{
    public class MistralClient : IDisposable
    {
        private readonly HttpClient _httpClient;
        private readonly string _baseUrl;
        private readonly string _apiKey;
        private bool _disposed;

        static MistralClient()
        {
            try
            {
                // Force TLS 1.2 (3072) and TLS 1.3 (12288) support in .NET 4.x
                ServicePointManager.SecurityProtocol |= (SecurityProtocolType)3072 | (SecurityProtocolType)12288 | SecurityProtocolType.Tls12;
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("Failed to set SecurityProtocol: {0}", ex.Message));
            }
        }

        public MistralClient(string baseUrl, string apiKey)
        {
            try
            {
                ServicePointManager.SecurityProtocol |= (SecurityProtocolType)3072 | (SecurityProtocolType)12288 | SecurityProtocolType.Tls12;
            }
            catch { }

            _baseUrl = !string.IsNullOrWhiteSpace(baseUrl) ? baseUrl.TrimEnd('/') : "https://api.mistral.ai/v1";
            _apiKey = apiKey ?? string.Empty;

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
            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        }

        public async Task<bool> TestConnectionAsync(CancellationToken ct = default(CancellationToken))
        {
            try
            {
                var testMessages = new List<ChatMessage>
                {
                    new ChatMessage("user", "Hello")
                };

                var requestPayload = new ChatRequest
                {
                    Model = "mistral-small-latest",
                    Messages = testMessages,
                    MaxTokens = 5,
                    Stream = false
                };

                string url = string.Format("{0}/chat/completions", _baseUrl);
                string json = JsonConvert.SerializeObject(requestPayload);
                using (var content = new StringContent(json, Encoding.UTF8, "application/json"))
                {
                    using (var response = await _httpClient.PostAsync(url, content, ct).ConfigureAwait(false))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            return true;
                        }

                        string errBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                        Logger.Warn(string.Format("TestConnection failed with HTTP {0}: {1}", (int)response.StatusCode, errBody));
                        throw new InvalidOperationException(string.Format("HTTP {0} ({1}): {2}", (int)response.StatusCode, response.ReasonPhrase, errBody));
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("TestConnection exception", ex);
                throw;
            }
        }

        public async Task<string> ChatAsync(
            string model,
            List<ChatMessage> messages,
            double temperature,
            int maxTokens,
            CancellationToken ct = default(CancellationToken))
        {
            string url = string.Format("{0}/chat/completions", _baseUrl);

            var requestPayload = new ChatRequest
            {
                Model = model,
                Messages = messages,
                Temperature = temperature,
                MaxTokens = maxTokens,
                Stream = false
            };

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
                        Logger.Warn(string.Format("Mistral API Error (Attempt {0}/{1}) HTTP {2}: {3}", attempt, maxRetries, statusCode, errorBody));

                        if ((statusCode == 429 || statusCode >= 500) && attempt < maxRetries)
                        {
                            shouldRetry = true;
                        }
                        else
                        {
                            throw new HttpRequestException(string.Format("Mistral API returned error {0} ({1}): {2}", (int)response.StatusCode, response.ReasonPhrase, errorBody));
                        }
                    }
                }
                catch (HttpRequestException)
                {
                    if (attempt < maxRetries)
                    {
                        shouldRetry = true;
                    }
                    else
                    {
                        throw;
                    }
                }

                if (shouldRetry)
                {
                    await Task.Delay(delayMs, ct).ConfigureAwait(false);
                    delayMs *= 2;
                }
            }

            throw new InvalidOperationException("Failed to get response from Mistral API after multiple attempts.");
        }

        public async Task StreamChatCallbackAsync(
            string model,
            List<ChatMessage> messages,
            double temperature,
            int maxTokens,
            Action<string> onDeltaReceived,
            CancellationToken ct = default(CancellationToken))
        {
            if (onDeltaReceived == null) throw new ArgumentNullException("onDeltaReceived");

            string url = string.Format("{0}/chat/completions", _baseUrl);

            var requestPayload = new ChatRequest
            {
                Model = model,
                Messages = messages,
                Temperature = temperature,
                MaxTokens = maxTokens,
                Stream = true
            };

            string requestJson = JsonConvert.SerializeObject(requestPayload);

            int maxRetries = 3;
            int delayMs = 1000;

            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                ct.ThrowIfCancellationRequested();
                bool shouldRetry = false;

                try
                {
                    var request = new HttpRequestMessage(HttpMethod.Post, url)
                    {
                        Content = new StringContent(requestJson, Encoding.UTF8, "application/json")
                    };

                    using (var response = await _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false))
                    {
                        if (!response.IsSuccessStatusCode)
                        {
                            int statusCode = (int)response.StatusCode;
                            string errorBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                            Logger.Warn(string.Format("Mistral API Streaming Error (Attempt {0}/{1}) HTTP {2}: {3}", attempt, maxRetries, statusCode, errorBody));

                            if ((statusCode == 429 || statusCode >= 500) && attempt < maxRetries)
                            {
                                shouldRetry = true;
                            }
                            else
                            {
                                throw new HttpRequestException(string.Format("Mistral API returned error {0} ({1}): {2}", (int)response.StatusCode, response.ReasonPhrase, errorBody));
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
                                        if (isDone) break;
                                        if (!string.IsNullOrEmpty(delta))
                                        {
                                            onDeltaReceived(delta);
                                        }
                                    }
                                }
                            }
                            return;
                        }
                    }
                }
                catch (HttpRequestException)
                {
                    if (attempt < maxRetries)
                    {
                        shouldRetry = true;
                    }
                    else
                    {
                        throw;
                    }
                }

                if (shouldRetry)
                {
                    await Task.Delay(delayMs, ct).ConfigureAwait(false);
                    delayMs *= 2;
                }
            }
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
