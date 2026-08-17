using System;
using Newtonsoft.Json.Linq;

namespace MistralOfficeAddin.Providers
{
    public static class GeminiStreamingParser
    {
        public static bool TryParseLine(string line, out string delta, out bool isDone)
        {
            delta = string.Empty;
            isDone = false;

            if (string.IsNullOrWhiteSpace(line)) return false;

            line = line.Trim();
            if (line.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
            {
                line = line.Substring(5).Trim();
            }

            if (string.IsNullOrEmpty(line) || line == "[DONE]")
            {
                isDone = (line == "[DONE]");
                return isDone;
            }

            try
            {
                var jObj = JObject.Parse(line);
                var candidates = jObj["candidates"] as JArray;
                if (candidates != null && candidates.Count > 0)
                {
                    var firstCandidate = candidates[0];
                    var finishReason = (string)firstCandidate["finishReason"];
                    if (!string.IsNullOrEmpty(finishReason) && finishReason != "MAX_TOKENS" && finishReason != "STOP")
                    {
                        // Some finish reasons like SAFETY etc.
                    }

                    var content = firstCandidate["content"];
                    if (content != null)
                    {
                        var parts = content["parts"] as JArray;
                        if (parts != null)
                        {
                            foreach (var part in parts)
                            {
                                var text = (string)part["text"];
                                if (!string.IsNullOrEmpty(text))
                                {
                                    delta += text;
                                }
                            }
                        }
                    }

                    if (!string.IsNullOrEmpty(finishReason) && finishReason == "STOP")
                    {
                        isDone = true;
                    }

                    return !string.IsNullOrEmpty(delta) || isDone;
                }
            }
            catch
            {
                // Incomplete or non-JSON SSE heartbeat line
            }

            return false;
        }
    }
}
