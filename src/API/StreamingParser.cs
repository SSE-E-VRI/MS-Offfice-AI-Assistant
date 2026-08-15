using System;
using MistralOfficeAddin.API.Models;
using MistralOfficeAddin.Core;
using Newtonsoft.Json;

namespace MistralOfficeAddin.API
{
    public static class StreamingParser
    {
        public const string DoneSignal = "[DONE]";

        public static bool TryParseLine(string line, out string deltaContent, out bool isDone)
        {
            deltaContent = null;
            isDone = false;

            if (string.IsNullOrWhiteSpace(line))
            {
                return false;
            }

            line = line.Trim();

            if (!line.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            string dataPayload = line.Substring(5).Trim();
            if (string.Equals(dataPayload, DoneSignal, StringComparison.OrdinalIgnoreCase))
            {
                isDone = true;
                return true;
            }

            try
            {
                var chunk = JsonConvert.DeserializeObject<StreamingChunk>(dataPayload);
                if (chunk != null && chunk.Choices != null && chunk.Choices.Count > 0)
                {
                    var choice = chunk.Choices[0];
                    if (choice != null && choice.Delta != null && !string.IsNullOrEmpty(choice.Delta.Content))
                    {
                        deltaContent = choice.Delta.Content;
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("Failed to parse SSE JSON payload: {0} -> Payload: {1}", ex.Message, dataPayload));
            }

            return false;
        }
    }
}
