using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using MistralOfficeAddin.API.Models;

namespace MistralOfficeAddin.Providers
{
    public enum AIProviderType
    {
        Mistral = 0,
        Groq = 1,
        Gemini = 2,
        Custom = 3
    }

    [Flags]
    public enum AICapabilities
    {
        None = 0,
        Chat = 1,
        Streaming = 2,
        Vision = 4,
        ModelListing = 8,
        ConnectionTest = 16
    }

    public class AIModelInfo
    {
        public string Id { get; set; }
        public string DisplayName { get; set; }
        public bool SupportsVision { get; set; }
        public bool SupportsStreaming { get; set; }

        public AIModelInfo()
        {
            SupportsStreaming = true;
        }

        public AIModelInfo(string id, string displayName = null, bool supportsVision = false)
        {
            Id = id;
            DisplayName = !string.IsNullOrWhiteSpace(displayName) ? displayName : id;
            SupportsVision = supportsVision;
            SupportsStreaming = true;
        }

        public override string ToString()
        {
            return DisplayName ?? Id ?? string.Empty;
        }
    }

    public class AttachmentBlock
    {
        public string FileName { get; set; }
        public string ContentType { get; set; }
        public string ExtractedText { get; set; }
        public byte[] RawBytes { get; set; }
        public bool IsImage { get; set; }
        public long FileSizeBytes { get; set; }
    }

    public class AIRequest
    {
        public string Model { get; set; }
        public List<ChatMessage> Messages { get; set; }
        public double Temperature { get; set; }
        public int MaxTokens { get; set; }
        public string SystemPrompt { get; set; }
        public List<AttachmentBlock> Attachments { get; set; }

        public AIRequest()
        {
            Messages = new List<ChatMessage>();
            Attachments = new List<AttachmentBlock>();
            Temperature = 0.7;
            MaxTokens = 4096;
        }
    }

    public class AIResponse
    {
        public string Content { get; set; }
        public string Model { get; set; }
        public int PromptTokens { get; set; }
        public int CompletionTokens { get; set; }

        public AIResponse()
        {
            Content = string.Empty;
        }

        public AIResponse(string content)
        {
            Content = content ?? string.Empty;
        }
    }

    public class AIException : Exception
    {
        public int StatusCode { get; private set; }
        public AIProviderType ProviderType { get; private set; }

        public AIException(string message, AIProviderType providerType, int statusCode = 0, Exception innerException = null)
            : base(message, innerException)
        {
            ProviderType = providerType;
            StatusCode = statusCode;
        }
    }

    public interface IAIProvider : IDisposable
    {
        AIProviderType ProviderType { get; }
        AICapabilities Capabilities { get; }

        Task<bool> TestConnectionAsync(CancellationToken ct = default(CancellationToken));
        Task<List<AIModelInfo>> ListModelsAsync(CancellationToken ct = default(CancellationToken));
        Task<AIResponse> ChatAsync(AIRequest request, CancellationToken ct = default(CancellationToken));
        Task StreamChatAsync(AIRequest request, Action<string> onDeltaReceived, CancellationToken ct = default(CancellationToken));
        bool CheckVisionSupport(string model);
    }
}
