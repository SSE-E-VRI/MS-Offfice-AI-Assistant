using System.Collections.Generic;
using Newtonsoft.Json;

namespace MSOfficeAIAssistant.API.Models
{
    public class ChatResponse
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
        public List<ChatChoice> Choices { get; set; }

        [JsonProperty("usage")]
        public UsageInfo Usage { get; set; }

        public ChatResponse()
        {
            Choices = new List<ChatChoice>();
        }
    }

    public class ChatChoice
    {
        [JsonProperty("index")]
        public int Index { get; set; }

        [JsonProperty("message")]
        public ChatMessage Message { get; set; }

        [JsonProperty("finish_reason")]
        public string FinishReason { get; set; }
    }

    public class UsageInfo
    {
        [JsonProperty("prompt_tokens")]
        public int PromptTokens { get; set; }

        [JsonProperty("completion_tokens")]
        public int CompletionTokens { get; set; }

        [JsonProperty("total_tokens")]
        public int TotalTokens { get; set; }
    }
}
