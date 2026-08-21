using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MSOfficeAIAssistant.API;
using MSOfficeAIAssistant.API.Models;
using MSOfficeAIAssistant.Attachments;
using MSOfficeAIAssistant.Core.Actions;
using MSOfficeAIAssistant.Providers;

namespace MSOfficeAIAssistant.Core.Session
{
    public class PreparedPayload
    {
        public AIRequest Request { get; set; }
        public string EffectiveSystemPrompt { get; set; }
        public List<AttachmentBlock> ExtractedAttachments { get; set; }
        public int DroppedImagesCount { get; set; }
    }

    /// <summary>
    /// Non-UI session coordinator managing conversation history, token budgeting,
    /// attachment processing, prompt preparation, response parsing, and persistence.
    /// Decoupled from WPF UI and testable headlessly.
    /// </summary>
    public class AssistantSession
    {
        private readonly ObservableCollection<ChatMessage> _messages;
        private readonly ChatOrchestrator _orchestrator;
        private readonly StreamCoordinator _streamCoordinator;
        private string _hostType;
        private string _currentDocumentKey = "OfficeSession";

        public ObservableCollection<ChatMessage> Messages
        {
            get { return _messages; }
        }

        public ChatOrchestrator Orchestrator
        {
            get { return _orchestrator; }
        }

        public StreamCoordinator StreamCoordinator
        {
            get { return _streamCoordinator; }
        }

        public string HostType
        {
            get { return _hostType; }
            set { _hostType = value; }
        }

        public string CurrentDocumentKey
        {
            get { return _currentDocumentKey; }
            set { _currentDocumentKey = value; }
        }

        public AssistantSession(ChatOrchestrator orchestrator = null, ObservableCollection<ChatMessage> messages = null)
        {
            _orchestrator = orchestrator ?? new ChatOrchestrator();
            _messages = messages ?? new ObservableCollection<ChatMessage>();
            _streamCoordinator = new StreamCoordinator();
        }

        public async Task<PreparedPayload> PreparePayloadAsync(
            string selectedModel,
            IEnumerable<string> attachmentFilePaths,
            ChatMessage assistantMsgToExclude = null)
        {
            var config = ConfigManager.Instance;
            var extractedAttachments = new List<AttachmentBlock>();
            var textAttachmentContext = new StringBuilder();
            int droppedImagesCount = 0;

            bool providerSupportsVision = _orchestrator != null && _orchestrator.CheckVisionSupport(selectedModel);

            if (attachmentFilePaths != null)
            {
                foreach (var path in attachmentFilePaths)
                {
                    if (string.IsNullOrWhiteSpace(path)) continue;
                    try
                    {
                        var block = await AttachmentExtractor.ExtractAsync(path);
                        if (block.IsImage)
                        {
                            if (providerSupportsVision)
                            {
                                extractedAttachments.Add(block);
                            }
                            else
                            {
                                droppedImagesCount++;
                            }
                        }
                        else
                        {
                            textAttachmentContext.AppendLine(string.Format(
                                "\n[Source: {0}]\n{1}\n[End Source: {0}]",
                                block.FileName,
                                block.ExtractedText));
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Warn(string.Format("Failed to extract attachment '{0}': {1}", path, ex.Message));
                    }
                }
            }

            // Clone message list for API request
            var historyForApi = _messages
                .Where(m => (m.IsUser || m.IsAssistant) && m != assistantMsgToExclude)
                .Select(m => new ChatMessage(m.Role, !string.IsNullOrEmpty(m.FullContent) ? m.FullContent : m.Content))
                .ToList();

            if (textAttachmentContext.Length > 0 && historyForApi.Count > 0)
            {
                var lastUser = historyForApi.LastOrDefault(m => m.IsUser);
                if (lastUser != null)
                {
                    lastUser.Content = PromptAssembler.AppendAttachmentCitationInstruction(lastUser.Content, textAttachmentContext.ToString());
                }
            }

            string effectiveSystemPrompt = PromptAssembler.BuildHostAwareSystemPrompt(config.SystemPrompt, _hostType);
            var boundedMessages = TokenCounter.TruncateToFit(historyForApi, 24000, effectiveSystemPrompt);

            var aiRequest = new AIRequest
            {
                Model = selectedModel,
                Messages = boundedMessages,
                Temperature = config.Temperature,
                MaxTokens = config.MaxTokens,
                SystemPrompt = effectiveSystemPrompt,
                Attachments = extractedAttachments
            };

            return new PreparedPayload
            {
                Request = aiRequest,
                EffectiveSystemPrompt = effectiveSystemPrompt,
                ExtractedAttachments = extractedAttachments,
                DroppedImagesCount = droppedImagesCount
            };
        }

        public void ProcessAssistantResponse(string fullAssistantText, ChatMessage assistantMsg)
        {
            if (assistantMsg == null) return;
            assistantMsg.IsStreaming = false;

            var extraction = ActionExtractor.Extract(fullAssistantText, _hostType);
            if (extraction != null && extraction.HasActions)
            {
                assistantMsg.Content = extraction.CleanText;
                foreach (var act in extraction.Actions)
                {
                    assistantMsg.OfficeActions.Add(act);
                }
                assistantMsg.NotifyOfficeActionsChanged();
                return;
            }

            assistantMsg.Content = fullAssistantText;
        }

        public void SaveHistory()
        {
            try
            {
                var list = _messages.Where(m => !m.IsSystem).ToList();
                ConversationStore.Instance.SaveHistory(_currentDocumentKey, list);
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("SaveConversationHistory failed: {0}", ex.Message));
            }
        }

        public void LoadHistory(string documentKey)
        {
            _currentDocumentKey = documentKey ?? "OfficeSession";
            try
            {
                var saved = ConversationStore.Instance.GetHistory(_currentDocumentKey);
                _messages.Clear();
                if (saved != null && saved.Count > 0)
                {
                    foreach (var msg in saved)
                    {
                        _messages.Add(msg);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("LoadConversationHistory failed: {0}", ex.Message));
            }
        }
    }
}
