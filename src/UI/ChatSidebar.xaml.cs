using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using MistralOfficeAddin.API;
using MistralOfficeAddin.API.Models;
using MistralOfficeAddin.Attachments;
using MistralOfficeAddin.Core;
using MistralOfficeAddin.Hosts;
using MistralOfficeAddin.Providers;

namespace MistralOfficeAddin.UI
{
    public class AttachmentItemViewModel
    {
        public string FilePath { get; set; }
        public string FileName { get; set; }
        public long FileSizeBytes { get; set; }
        public bool IsImage { get; set; }

        public string DisplayIcon
        {
            get { return IsImage ? "🖼️" : "📄"; }
        }

        public string DisplayName
        {
            get { return string.Format("{0} ({1:F1} KB)", FileName, FileSizeBytes / 1024.0); }
        }
    }

    public partial class ChatSidebar : UserControl
    {
        private readonly ObservableCollection<ChatMessage> _messages = new ObservableCollection<ChatMessage>();
        private readonly ObservableCollection<AttachmentItemViewModel> _pendingAttachments = new ObservableCollection<AttachmentItemViewModel>();
        private readonly ChatOrchestrator _orchestrator;
        private CancellationTokenSource _streamingCts;
        private string _currentDocumentKey = "OfficeSession";
        private object _hostAppObj;
        private string _hostType = "Office";
        private bool _isSending = false;

        // Host Controllers — created lazily, not during startup
        private WordController _wordCtrl;
        private ExcelController _excelCtrl;
        private PowerPointController _pptCtrl;
        private bool _hostInitialized;

        public ChatSidebar()
        {
            InitializeComponent();
            MessagesItemsControl.ItemsSource = _messages;
            AttachmentsItemsControl.ItemsSource = _pendingAttachments;

            _orchestrator = new ChatOrchestrator(ProviderFactory.CreateFromConfig(ConfigManager.Instance));

            SelectConfiguredModel();

            // Show welcome message immediately without waiting for host init
            if (ConfigManager.Instance.LoadFailed)
            {
                _messages.Add(new ChatMessage("system", "⚠ Configuration could not be read and was reset to defaults. Please re-enter your settings in Configure (⚙️)."));
            }
            else
            {
                _messages.Add(new ChatMessage("system", "AI Assistant is ready. Click Configure (⚙️) to enter your API key, then start chatting!"));
            }
        }

        public void ReloadConfiguredProvider()
        {
            try
            {
                if (_orchestrator != null)
                {
                    var newProvider = ProviderFactory.CreateFromConfig(ConfigManager.Instance);
                    _orchestrator.UpdateProvider(newProvider);
                }
                SelectConfiguredModel();
            }
            catch (Exception ex)
            {
                Logger.Error("ChatSidebar: Failed to reload provider", ex);
            }
        }

        private void SelectConfiguredModel()
        {
            try
            {
                var config = ConfigManager.Instance;
                var providerSettings = config.GetActiveProviderSettings();
                string configured = providerSettings.DefaultModel;

                if (CmbModel == null) return;

                CmbModel.Items.Clear();
                var defaultModels = GetDefaultModels(config.ActiveProvider);
                foreach (var m in defaultModels)
                {
                    CmbModel.Items.Add(m.Id);
                }

                if (!string.IsNullOrWhiteSpace(configured))
                {
                    if (!CmbModel.Items.Contains(configured))
                    {
                        CmbModel.Items.Insert(0, configured);
                    }
                    CmbModel.SelectedItem = configured;
                }
                else if (CmbModel.Items.Count > 0)
                {
                    CmbModel.SelectedIndex = 0;
                }
            }
            catch { }
        }

        private List<AIModelInfo> GetDefaultModels(AIProviderType providerType)
        {
            switch (providerType)
            {
                case AIProviderType.Groq:
                    return new List<AIModelInfo>
                    {
                        new AIModelInfo("llama-3.3-70b-versatile"),
                        new AIModelInfo("llama-3.1-8b-instant"),
                        new AIModelInfo("llama-3.2-11b-vision-preview"),
                        new AIModelInfo("llama-3.2-90b-vision-preview"),
                        new AIModelInfo("mixtral-8x7b-32768")
                    };
                case AIProviderType.Gemini:
                    return new List<AIModelInfo>
                    {
                        new AIModelInfo("gemini-2.5-flash"),
                        new AIModelInfo("gemini-2.5-pro"),
                        new AIModelInfo("gemini-1.5-flash"),
                        new AIModelInfo("gemini-1.5-pro")
                    };
                case AIProviderType.Custom:
                    return new List<AIModelInfo>
                    {
                        new AIModelInfo("llama3"),
                        new AIModelInfo("mistral"),
                        new AIModelInfo("qwen2.5")
                    };
                case AIProviderType.Mistral:
                default:
                    return new List<AIModelInfo>
                    {
                        new AIModelInfo("mistral-large-latest"),
                        new AIModelInfo("mistral-small-latest"),
                        new AIModelInfo("open-mistral-nemo"),
                        new AIModelInfo("codestral-latest"),
                        new AIModelInfo("pixtral-large-latest")
                    };
            }
        }

        /// <summary>
        /// Called after the task pane is visible. Deferred host controller creation
        /// happens on a background thread to avoid blocking the COM thread.
        /// </summary>
        public void InitializeHost(object appObj, string hostType)
        {
            _hostAppObj = appObj;
            _hostType = hostType ?? "Office";

            // Update badge immediately (UI thread, no COM calls)
            TxtDocumentBadge.Text = string.Format("{0} Document", _hostType);

            // Host controllers must be created on the Office STA UI thread.
            Dispatcher.BeginInvoke(new Action(InitializeHostOnUiThread), DispatcherPriority.Background);
        }

        private void InitializeHostOnUiThread()
        {
            if (_hostInitialized) return;

            try
            {
                string docName = "Document";

                if (string.Equals(_hostType, "Word", StringComparison.OrdinalIgnoreCase))
                {
                    _wordCtrl = new WordController(_hostAppObj);
                    docName = _wordCtrl.GetActiveDocumentName();
                }
                else if (string.Equals(_hostType, "Excel", StringComparison.OrdinalIgnoreCase))
                {
                    _excelCtrl = new ExcelController(_hostAppObj);
                    docName = _excelCtrl.GetActiveWorkbookName();
                }
                else if (string.Equals(_hostType, "PowerPoint", StringComparison.OrdinalIgnoreCase))
                {
                    _pptCtrl = new PowerPointController(_hostAppObj);
                    docName = _pptCtrl.GetActivePresentationName();
                }

                _currentDocumentKey = string.IsNullOrWhiteSpace(docName) ? "Document" : docName;
                _hostInitialized = true;
                TxtDocumentBadge.Text = string.Format("{0}: {1}", _hostType, _currentDocumentKey);
                LoadConversationHistory();
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("ChatSidebar.InitializeHostOnUiThread failed: {0}", ex.Message));
            }
        }

        private void LoadConversationHistory()
        {
            _messages.Clear();
            try
            {
                var history = ConversationStore.Instance.GetHistory(_currentDocumentKey);
                if (history != null && history.Count > 0)
                {
                    foreach (var msg in history)
                    {
                        _messages.Add(msg);
                    }
                }
                else
                {
                    _messages.Add(new ChatMessage("system", string.Format("Welcome to AI Assistant for {0}! Ask anything or use ribbon buttons to draft, rewrite, or summarize.", _hostType)));
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("LoadConversationHistory failed: {0}", ex.Message));
            }
            ScrollToBottom();
        }

        private void SaveConversationHistory()
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

        public async void ExecuteExternalPrompt(string prompt, string promptTitle)
        {
            if (string.IsNullOrWhiteSpace(prompt)) return;

            if (_isSending && _streamingCts != null)
            {
                try { _streamingCts.Cancel(); } catch { }
            }

            string selectedText = GetSelectedTextOnly();
            string fullPrompt;
            string displayTitle = promptTitle;

            if (!string.IsNullOrWhiteSpace(selectedText))
            {
                if (promptTitle != null && promptTitle.StartsWith("Translate", StringComparison.OrdinalIgnoreCase))
                {
                    fullPrompt = string.Format("{0}\n\nText to translate:\n\"\"\"\n{1}\n\"\"\"", prompt, selectedText);
                }
                else
                {
                    fullPrompt = string.Format("{0}\n\n[Selected Text]:\n{1}", prompt, selectedText);
                }

                string snippet = selectedText.Trim().Replace('\r', ' ').Replace('\n', ' ');
                if (snippet.Length > 80) snippet = snippet.Substring(0, 77) + "...";
                displayTitle = string.Format("{0}:\n\"{1}\"", promptTitle, snippet);
            }
            else
            {
                string contextText = GetCurrentContextText(false);
                fullPrompt = string.IsNullOrEmpty(contextText)
                    ? prompt
                    : string.Format("{0}\n\n[Context]:\n{1}", prompt, contextText);
            }

            await SendMessageAsync(fullPrompt, displayTitle);
        }

        private string GetSelectedTextOnly()
        {
            try
            {
                if (_wordCtrl != null)
                {
                    string sel = _wordCtrl.GetSelectedText();
                    if (!string.IsNullOrWhiteSpace(sel)) return sel;
                }
                if (_excelCtrl != null)
                {
                    string sel = _excelCtrl.GetSelectedRangeValues();
                    if (!string.IsNullOrWhiteSpace(sel)) return sel;
                }
                if (_pptCtrl != null)
                {
                    string sel = _pptCtrl.GetSlideText();
                    if (!string.IsNullOrWhiteSpace(sel)) return sel;
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("GetSelectedTextOnly failed: {0}", ex.Message));
            }
            return string.Empty;
        }

        private async void BtnSend_Click(object sender, RoutedEventArgs e)
        {
            if (_isSending) return;
            string text = TxtInput.Text.Trim();
            if (string.IsNullOrWhiteSpace(text)) return;
            TxtInput.Clear();

            string contextText = (ChkIncludeSelection.IsChecked == true) ? GetCurrentContextText(true) : string.Empty;
            string fullPrompt = string.IsNullOrEmpty(contextText)
                ? text
                : string.Format("{0}\n\n[Context]:\n{1}", text, contextText);

            await SendMessageAsync(fullPrompt, null);
        }

        private async Task SendMessageAsync(string promptToSend, string customDisplayTitle)
        {
            var config = ConfigManager.Instance;
            if (string.IsNullOrWhiteSpace(config.ApiKey) && config.ActiveProvider != AIProviderType.Custom)
            {
                _messages.Add(new ChatMessage("system", "API Key is missing. Click ⚙️ Settings to configure your AI provider key."));
                ScrollToBottom();
                return;
            }

            if (_isSending && _streamingCts != null)
            {
                try { _streamingCts.Cancel(); } catch { }
            }

            _isSending = true;
            string displayUserMessage = customDisplayTitle ?? promptToSend;
            _messages.Add(new ChatMessage("user", displayUserMessage));
            ScrollToBottom();

            var assistantMsg = new ChatMessage("assistant", "") { IsStreaming = true };
            _messages.Add(assistantMsg);
            ScrollToBottom();

            BtnSend.IsEnabled = false;
            TypingIndicator.Visibility = Visibility.Visible;

            if (_streamingCts != null)
            {
                try { _streamingCts.Cancel(); } catch { }
                try { _streamingCts.Dispose(); } catch { }
            }
            _streamingCts = new CancellationTokenSource();
            var myStreamCts = _streamingCts;

            string selectedModel = !string.IsNullOrWhiteSpace(CmbModel.Text)
                ? CmbModel.Text.Trim()
                : (config.DefaultModel ?? "default");

            try
            {
                // Process pending attachments without leaving UI thread
                var extractedAttachments = new List<AttachmentBlock>();
                var textAttachmentContext = new StringBuilder();

                if (_pendingAttachments.Count > 0)
                {
                    bool providerSupportsVision = _orchestrator != null && _orchestrator.CheckVisionSupport(selectedModel);
                    int droppedImagesCount = 0;

                    foreach (var att in _pendingAttachments)
                    {
                        try
                        {
                            var block = await AttachmentExtractor.ExtractAsync(att.FilePath);
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
                                textAttachmentContext.AppendLine(string.Format("\n[Attachment: {0}]\n{1}\n[End Attachment]", block.FileName, block.ExtractedText));
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.Warn(string.Format("Failed to extract attachment '{0}': {1}", att.FileName, ex.Message));
                        }
                    }

                    if (droppedImagesCount > 0)
                    {
                        _messages.Insert(_messages.IndexOf(assistantMsg), new ChatMessage("system", "⚠ Note: Image attachments were omitted — the selected model does not support vision analysis."));
                    }
                }

                // Clone message list for API request to avoid mutating bound UI user messages
                var historyForApi = _messages
                    .Where(m => (m.IsUser || m.IsAssistant) && m != assistantMsg)
                    .Select(m => new ChatMessage(m.Role, m.Content))
                    .ToList();

                // If attachments produced extracted text, augment the last user message copy
                if (textAttachmentContext.Length > 0 && historyForApi.Count > 0)
                {
                    var lastUser = historyForApi.LastOrDefault(m => m.IsUser);
                    if (lastUser != null)
                    {
                        lastUser.Content = lastUser.Content + "\n\n" + textAttachmentContext.ToString();
                    }
                }

                string effectiveSystemPrompt = BuildHostAwareSystemPrompt(config.SystemPrompt);
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

                var tokenAccumulator = new StringBuilder();

                await _orchestrator.StreamChatAsync(
                    aiRequest,
                    delta =>
                    {
                        if (string.IsNullOrEmpty(delta)) return;
                        lock (tokenAccumulator)
                        {
                            tokenAccumulator.Append(delta);
                        }
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            assistantMsg.Content += delta;
                        }), DispatcherPriority.Background);
                    },
                    myStreamCts.Token);

                string fullAssistantText;
                lock (tokenAccumulator)
                {
                    fullAssistantText = tokenAccumulator.ToString();
                }

                Dispatcher.Invoke(new Action(() =>
                {
                    assistantMsg.Content = fullAssistantText;
                    assistantMsg.IsStreaming = false;

                    // Parse structured spreadsheet actions from Excel responses
                    string cleanContent;
                    var extractedActions = SpreadsheetActionParser.ExtractActions(fullAssistantText, out cleanContent);
                    if (extractedActions != null && extractedActions.Count > 0)
                    {
                        assistantMsg.Content = cleanContent;
                        foreach (var act in extractedActions)
                        {
                            assistantMsg.Actions.Add(act);
                        }
                        assistantMsg.NotifyActionsChanged();
                    }

                    // Scroll once at the end, not on every token
                    ScrollToBottom();
                    SaveConversationHistory();

                    // Clear attachments on UI thread after send
                    _pendingAttachments.Clear();
                    UpdateAttachmentState();
                }));
            }
            catch (OperationCanceledException)
            {
                Dispatcher.Invoke(new Action(() =>
                {
                    assistantMsg.Content += "\n\n*(Generation stopped)*";
                    assistantMsg.IsStreaming = false;
                }));
            }
            catch (Exception ex)
            {
                Logger.Error("Chat completion error", ex);
                Dispatcher.Invoke(new Action(() =>
                {
                    assistantMsg.Content = string.Format("Error: {0}", ex.Message);
                    assistantMsg.IsStreaming = false;
                }));
            }
            finally
            {
                _isSending = false;
                Dispatcher.Invoke(new Action(() =>
                {
                    BtnSend.IsEnabled = true;
                    TypingIndicator.Visibility = Visibility.Collapsed;
                    ScrollToBottom();
                }));
            }
        }

        private void BtnAttach_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new Microsoft.Win32.OpenFileDialog
                {
                    Title = "Attach Documents or Images",
                    Multiselect = true,
                    Filter = "All Supported (*.docx;*.xlsx;*.pptx;*.pdf;*.png;*.jpg;*.txt)|*.docx;*.xlsx;*.pptx;*.pdf;*.png;*.jpg;*.jpeg;*.webp;*.gif;*.txt;*.csv;*.json;*.md|" +
                             "Office Documents (*.docx;*.xlsx;*.pptx)|*.docx;*.xlsx;*.pptx|" +
                             "PDF Documents (*.pdf)|*.pdf|" +
                             "Images (*.png;*.jpg;*.jpeg;*.webp)|*.png;*.jpg;*.jpeg;*.webp;*.gif|" +
                             "Text Files (*.txt;*.csv;*.json;*.md)|*.txt;*.csv;*.json;*.md|" +
                             "All Files (*.*)|*.*"
                };

                if (dlg.ShowDialog() == true && dlg.FileNames != null)
                {
                    if (_pendingAttachments.Count + dlg.FileNames.Length > AttachmentExtractor.MaxFileCount)
                    {
                        MessageBox.Show(
                            string.Format("You cannot attach more than {0} files per message.", AttachmentExtractor.MaxFileCount),
                            "AI Assistant - Attachment Limit",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                        return;
                    }

                    long currentTotal = _pendingAttachments.Sum(a => a.FileSizeBytes);

                    foreach (string file in dlg.FileNames)
                    {
                        var fi = new FileInfo(file);
                        if (!fi.Exists) continue;

                        string ext = fi.Extension.ToLowerInvariant();
                        if (ext == ".doc" || ext == ".xls" || ext == ".ppt" || ext == ".rtf")
                        {
                            MessageBox.Show(
                                string.Format("Legacy binary format '{0}' is not supported.\nPlease save as modern Open XML (.docx, .xlsx, .pptx) or export to PDF.", fi.Name),
                                "AI Assistant - Unsupported Format",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning);
                            continue;
                        }

                        if (fi.Length > AttachmentExtractor.MaxPerFileSizeBytes)
                        {
                            MessageBox.Show(
                                string.Format("File '{0}' exceeds the maximum allowed single file size of 20 MB.", fi.Name),
                                "AI Assistant - File Too Large",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning);
                            continue;
                        }

                        if (currentTotal + fi.Length > AttachmentExtractor.MaxTotalSizeBytes)
                        {
                            MessageBox.Show(
                                "Total attachments exceed the aggregate 30 MB size limit.",
                                "AI Assistant - Attachment Limit",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning);
                            break;
                        }

                        bool isImg = ext == ".png" || ext == ".jpg" || ext == ".jpeg" || ext == ".webp" || ext == ".gif" || ext == ".bmp";
                        _pendingAttachments.Add(new AttachmentItemViewModel
                        {
                            FilePath = fi.FullName,
                            FileName = fi.Name,
                            FileSizeBytes = fi.Length,
                            IsImage = isImg
                        });
                        currentTotal += fi.Length;
                    }

                    UpdateAttachmentState();
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Error attaching files", ex);
                MessageBox.Show("Could not attach file: " + ex.Message, "AI Assistant", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnRemoveAttachment_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            if (btn != null)
            {
                var item = btn.DataContext as AttachmentItemViewModel;
                if (item != null)
                {
                    _pendingAttachments.Remove(item);
                    UpdateAttachmentState();
                }
            }
        }

        private void UpdateAttachmentState()
        {
            AttachmentsItemsControl.Visibility = _pendingAttachments.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
            UpdateVisionWarning();
        }

        private void UpdateVisionWarning()
        {
            bool hasImages = _pendingAttachments.Any(a => a.IsImage);
            string selectedModel = !string.IsNullOrWhiteSpace(CmbModel.Text) ? CmbModel.Text.Trim() : (ConfigManager.Instance.DefaultModel ?? "");
            bool supportsVision = _orchestrator != null && _orchestrator.CheckVisionSupport(selectedModel);

            VisionWarningBanner.Visibility = (hasImages && !supportsVision) ? Visibility.Visible : Visibility.Collapsed;
        }

        private string BuildHostAwareSystemPrompt(string basePrompt)
        {
            string hostContext = "";
            if (_hostType == "Excel" || _excelCtrl != null)
            {
                hostContext = "\n\nYou are embedded inside Microsoft Excel.\n" +
                    "When generating calculations, formulas, or spreadsheet changes:\n" +
                    "1. Inspect the provided Worksheet Context with its explicit Column Letters (Col A, Col B, Col C, etc.) and Header names.\n" +
                    "2. When matching category/text columns with prefixes (e.g. '0001-non ferrous items'), use wildcard criteria (e.g. \"*non ferrous*\") in SUMIF/COUNTIF.\n" +
                    "3. Start row-level data formulas at Row 2 (e.g. F2, G2) rather than header Row 1.\n" +
                    "4. ALWAYS return executable spreadsheet actions in a structured <excel_actions> XML block:\n" +
                    "   <excel_actions>\n" +
                    "     <excel_action target=\"K20\" type=\"formula\" formula=\"=SUMIF(B:B, &quot;*non ferrous*&quot;, F:F)\" description=\"Total non-ferrous value\" />\n" +
                    "     <excel_action target=\"K21\" type=\"formula\" formula=\"=COUNTIF(E:E, 0)\" description=\"Count of zero-quantity items\" />\n" +
                    "     <excel_action target=\"K22\" type=\"formula\" formula=\"=AVERAGEIF(E:E, 0, F:F)\" description=\"Average value of zero-quantity items\" />\n" +
                    "     <excel_action target=\"G2:G27\" type=\"filldown\" formula=\"=IF(F2&gt;50000, &quot;High Value&quot;, &quot;&quot;)\" description=\"High value flag (&gt;50,000)\" />\n" +
                    "   </excel_actions>\n" +
                    "5. Provide a brief conversational summary above or below the action block without tutorial how-to steps.";
            }
            else if (_hostType == "Word" || _wordCtrl != null)
            {
                hostContext = "\n\nYou are embedded inside Microsoft Word. When the user asks to write, edit, rewrite, summarize, or translate text, provide the polished text directly without tutorial meta-commentary.";
            }
            else if (_hostType == "PowerPoint" || _pptCtrl != null)
            {
                hostContext = "\n\nYou are embedded inside Microsoft PowerPoint. When the user asks for slides or bullet points, provide structured slides with Slide titles, concise bullet points, and speaker notes.";
            }
            return (basePrompt ?? "You are an expert AI assistant embedded inside Microsoft Office.") + hostContext;
        }

        private string GetCurrentContextText(bool selectionOnly)
        {
            try
            {
                if (_wordCtrl != null)
                {
                    string sel = _wordCtrl.GetSelectedText();
                    if (!string.IsNullOrWhiteSpace(sel))
                    {
                        return string.Format("[Selected Text]:\n{0}", sel);
                    }
                    if (selectionOnly)
                    {
                        return string.Empty;
                    }
                    string docText = _wordCtrl.GetDocumentText(24000);
                    if (!string.IsNullOrWhiteSpace(docText))
                    {
                        return string.Format("[Document Content: {0}]:\n{1}", _currentDocumentKey, docText);
                    }
                }
                if (_excelCtrl != null)
                {
                    if (selectionOnly)
                    {
                        string sel = _excelCtrl.GetSelectedRangeValues();
                        if (!string.IsNullOrWhiteSpace(sel))
                        {
                            return sel;
                        }
                    }
                    string snapshot = _excelCtrl.GetWorksheetSnapshot(70, 26);
                    if (!string.IsNullOrWhiteSpace(snapshot))
                    {
                        return snapshot;
                    }
                }
                if (_pptCtrl != null) return _pptCtrl.GetSlideText();
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("GetCurrentContextText failed: {0}", ex.Message));
            }
            return string.Empty;
        }

        private void BtnStopStreaming_Click(object sender, RoutedEventArgs e)
        {
            if (_streamingCts != null)
            {
                _streamingCts.Cancel();
            }
            if (_orchestrator != null)
            {
                _orchestrator.CancelCurrentStream();
            }
        }

        public void StartNewChat(bool confirm)
        {
            if (confirm)
            {
                if (MessageBox.Show("Start a new chat session and clear current conversation?", "AI Assistant",
                    MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                {
                    return;
                }
            }

            _messages.Clear();
            ConversationStore.Instance.ClearHistory(_currentDocumentKey);
            _messages.Add(new ChatMessage("system", "New conversation started."));
        }

        private void BtnNewChat_Click(object sender, RoutedEventArgs e)
        {
            StartNewChat(true);
        }

        private void BtnSettings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var win = new SettingsWindow();
                if (win.ShowDialog() == true)
                {
                    ReloadConfiguredProvider();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Could not open settings: {0}", ex.Message), "AI Assistant", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private ChatMessage GetMessageFromSender(object sender)
        {
            var btn = sender as Button;
            if (btn == null) return null;
            return btn.DataContext as ChatMessage;
        }

        private void BtnApplyAction_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            var action = btn != null ? btn.Tag as SpreadsheetAction : null;
            if (action == null) return;

            if (_excelCtrl != null)
            {
                bool ok = _excelCtrl.ApplySpreadsheetAction(action);
                if (!ok && !string.IsNullOrEmpty(action.ErrorMessage))
                {
                    MessageBox.Show(action.ErrorMessage, "Spreadsheet Action Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            else
            {
                MessageBox.Show("No active Excel spreadsheet found.", "AI Assistant", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void BtnApplyAllActions_Click(object sender, RoutedEventArgs e)
        {
            var msg = GetMessageFromSender(sender);
            if (msg == null || msg.Actions == null || msg.Actions.Count == 0) return;

            if (_excelCtrl == null)
            {
                MessageBox.Show("No active Excel spreadsheet found.", "AI Assistant", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            int appliedCount = 0;
            foreach (var act in msg.Actions)
            {
                if (act.Status == SpreadsheetActionStatus.Pending || act.Status == SpreadsheetActionStatus.Error)
                {
                    if (_excelCtrl.ApplySpreadsheetAction(act))
                    {
                        appliedCount++;
                    }
                }
            }
        }

        private void BtnCopyMessage_Click(object sender, RoutedEventArgs e)
        {
            var msg = GetMessageFromSender(sender);
            if (msg == null || string.IsNullOrEmpty(msg.Content)) return;
            try { Clipboard.SetText(msg.Content); } catch { }
        }

        private void BtnInsertMessage_Click(object sender, RoutedEventArgs e)
        {
            var msg = GetMessageFromSender(sender);
            if (msg == null || string.IsNullOrEmpty(msg.Content)) return;
            string content = msg.Content;
            try
            {
                if (_wordCtrl != null)
                    _wordCtrl.InsertTextAtCursor(content);
                else if (_excelCtrl != null)
                    _excelCtrl.InsertText(content);
                else if (_pptCtrl != null)
                    _pptCtrl.InsertText(content);
                else
                    MessageBox.Show("No Office document is active. Open a document and try again.", "AI Assistant", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Could not insert text: {0}", ex.Message), "AI Assistant", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnQuickAction_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            if (btn != null && btn.Tag is string)
            {
                string prompt = (string)btn.Tag;
                string title = btn.Content != null ? btn.Content.ToString() : null;
                ExecuteExternalPrompt(prompt, title);
            }
        }

        private void TxtInput_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                TxtInput.Focus();
                Keyboard.Focus(TxtInput);
                try
                {
                    var source = System.Windows.PresentationSource.FromVisual(TxtInput) as System.Windows.Interop.HwndSource;
                    if (source != null && source.Handle != IntPtr.Zero)
                    {
                        IntPtr curFocus = NativeWnd.GetFocus();
                        if (curFocus != source.Handle && !NativeWnd.IsChild(source.Handle, curFocus))
                        {
                            NativeWnd.SetFocus(source.Handle);
                        }
                    }
                }
                catch { }
            }
        }

        private void TxtInput_GotFocus(object sender, RoutedEventArgs e)
        {
            try
            {
                var source = System.Windows.PresentationSource.FromVisual(TxtInput) as System.Windows.Interop.HwndSource;
                if (source != null && source.Handle != IntPtr.Zero)
                {
                    IntPtr curFocus = NativeWnd.GetFocus();
                    if (curFocus != source.Handle && !NativeWnd.IsChild(source.Handle, curFocus))
                    {
                        NativeWnd.SetFocus(source.Handle);
                    }
                }
            }
            catch { }
        }

        private void TxtInput_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
            {
                e.Handled = true;
                BtnSend_Click(sender, e);
            }
        }

        private void CmbModel_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_hostInitialized) return;
            string selected = (CmbModel != null ? CmbModel.SelectedItem as string : null) ?? (CmbModel != null ? CmbModel.Text : null);
            if (!string.IsNullOrWhiteSpace(selected))
            {
                ConfigManager.Instance.DefaultModel = selected.Trim();
                ConfigManager.Instance.Save();
            }
        }

        private void ScrollToBottom()
        {
            try
            {
                if (ChatScrollViewer != null)
                    ChatScrollViewer.ScrollToBottom();
            }
            catch { }
        }
    }
}
