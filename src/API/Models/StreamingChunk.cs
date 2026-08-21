using System.Collections.Generic;
using Newtonsoft.Json;

namespace MSOfficeAIAssistant.API.Models
{
    public class StreamingChunk
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("object")]
        public string Object { get; set; }

        [JsonProperty("created")]
        public long Created { get; set; }

        [JsonProperty("model")]
        public string Model { get; set; }

        [JsonProperty("choices")]
        public List<StreamingChoice> Choices { get; set; }

        public StreamingChunk()
        {
            Choices = new List<StreamingChoice>();
        }
    }

    public class StreamingChoice
    {
        [JsonProperty("index")]
        public int Index { get; set; }

        [JsonProperty("delta")]
        public DeltaContent Delta { get; set; }

        [JsonProperty("finish_reason")]
        public string FinishReason { get; set; }
    }

    public class DeltaContent
    {
        [JsonProperty("role")]
        public string Role { get; set; }

        [JsonProperty("content")]
        public string Content { get; set; }
    }
}
