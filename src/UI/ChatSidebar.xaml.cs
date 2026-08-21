using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Threading;
using MSOfficeAIAssistant.API;
using MSOfficeAIAssistant.API.Models;
using MSOfficeAIAssistant.Attachments;
using MSOfficeAIAssistant.Core;
using MSOfficeAIAssistant.Hosts;
using MSOfficeAIAssistant.Providers;

namespace MSOfficeAIAssistant.UI
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
        public event EventHandler PromptInputFocusRequested;
        public event EventHandler PromptInputFocusLost;

        private enum PromptContextScope
        {
            Selection,
            CurrentFile,
            SelectionAndFile,
            AttachmentsOnly
        }

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

            this.Loaded += ChatSidebar_Loaded;

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
                        new AIModelInfo("openai/gpt-oss-20b"),
                        new AIModelInfo("openai/gpt-oss-120b"),
                        new AIModelInfo("qwen/qwen3.6-27b")
                    };
                case AIProviderType.Gemini:
                    return new List<AIModelInfo>
                    {
                        new AIModelInfo("gemini-3.6-flash")
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
                UpdateHostSpecificControls();
                LoadConversationHistory();
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("ChatSidebar.InitializeHostOnUiThread failed: {0}", ex.Message));
            }
        }

        private void UpdateHostSpecificControls()
        {
            if (ChkTrackChanges == null) return;
            bool isWord = _wordCtrl != null;
            bool isPowerPoint = _pptCtrl != null;
            ChkTrackChanges.IsEnabled = isWord;
            if (BtnAcceptRevisions != null) BtnAcceptRevisions.Visibility = isWord ? Visibility.Visible : Visibility.Collapsed;
            if (BtnRejectRevisions != null) BtnRejectRevisions.Visibility = isWord ? Visibility.Visible : Visibility.Collapsed;
            if (BtnQuickDeck != null) BtnQuickDeck.Visibility = isPowerPoint ? Visibility.Visible : Visibility.Collapsed;
            if (BtnInsertVisual != null) BtnInsertVisual.Visibility = isPowerPoint ? Visibility.Visible : Visibility.Collapsed;
            if (!isWord)
            {
                ChkTrackChanges.IsChecked = false;
                ChkTrackChanges.ToolTip = "Track edits is available when Word is the active host.";
            }
            else
            {
                ChkTrackChanges.ToolTip = "Insert/rewrite the response as Word Track Changes.";
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
            await ExecuteExternalPromptAsync(prompt, promptTitle, GetPromptContextScope());
        }

        private async Task ExecuteExternalPromptAsync(string prompt, string promptTitle, PromptContextScope scope)
        {
            if (string.IsNullOrWhiteSpace(prompt)) return;

            // Cancel any in-flight stream and wait for isSending to clear
            if (_isSending)
            {
                if (_streamingCts != null)
                {
                    try { _streamingCts.Cancel(); } catch { }
                }
                // Brief yield to allow the in-flight finally block to run and reset _isSending
                await Task.Delay(100).ConfigureAwait(true);
            }
            // If still sending after grace period, skip — don't layer concurrent streams
            if (_isSending) return;

            string selectedText = IncludesSelection(scope) ? GetSelectedTextOnly() : string.Empty;
            string documentContext = IncludesCurrentFile(scope) ? GetCurrentFileContext(prompt) : string.Empty;
            string fullPrompt = prompt;
            string displayTitle = promptTitle;

            if (!string.IsNullOrWhiteSpace(selectedText))
            {
                if (promptTitle != null && promptTitle.StartsWith("Translate", StringComparison.OrdinalIgnoreCase))
                {
                    fullPrompt = string.Format("{0}\n\nText to translate:\n\"\"\"\n{1}\n\"\"\"", fullPrompt, selectedText);
                }
                else
                {
                    fullPrompt = string.Format("{0}\n\n[Selected Context]:\n{1}", fullPrompt, selectedText);
                }

                string snippet = selectedText.Trim().Replace('\r', ' ').Replace('\n', ' ');
                if (snippet.Length > 80) snippet = snippet.Substring(0, 77) + "...";
                displayTitle = string.Format("{0}:\n\"{1}\"", promptTitle, snippet);
            }
            if (!string.IsNullOrWhiteSpace(documentContext))
            {
                fullPrompt = string.Format("{0}\n\n[Current File Context]:\n{1}", fullPrompt, documentContext);
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

            string fullPrompt = ComposePromptWithContext(text, GetPromptContextScope());

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
            var userMsg = new ChatMessage("user", displayUserMessage);
            userMsg.FullContent = promptToSend;
            _messages.Add(userMsg);
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
                                textAttachmentContext.AppendLine(string.Format(
                                    "\n[Source: {0}]\n{1}\n[End Source: {0}]",
                                    block.FileName,
                                    block.ExtractedText));
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
                    .Select(m => new ChatMessage(m.Role, !string.IsNullOrEmpty(m.FullContent) ? m.FullContent : m.Content))
                    .ToList();

                // If attachments produced extracted text, augment the last user message copy
                if (textAttachmentContext.Length > 0 && historyForApi.Count > 0)
                {
                    var lastUser = historyForApi.LastOrDefault(m => m.IsUser);
                    if (lastUser != null)
                    {
                        lastUser.Content = lastUser.Content + "\n\n" + textAttachmentContext.ToString() +
                            "\nWhen you rely on an attached source, cite it using [Source: filename, page/section] and do not invent a source.";
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
                int _deltaCount = 0;
                bool streamFinished = false;

                await _orchestrator.StreamChatAsync(
                    aiRequest,
                    delta =>
                    {
                        if (string.IsNullOrEmpty(delta)) return;
                        lock (tokenAccumulator)
                        {
                            tokenAccumulator.Append(delta);
                            _deltaCount++;
                        }
                        // Throttled UI update: only update display every 5 deltas to reduce dispatcher pressure
                        if (_deltaCount % 5 == 0)
                        {
                            string snapshot;
                            lock (tokenAccumulator)
                            {
                                snapshot = tokenAccumulator.ToString();
                            }
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                // Ignore queued updates after the stream has completed so a stale
                                // snapshot never overwrites the final content.
                                if (!streamFinished)
                                {
                                    assistantMsg.Content = snapshot;
                                }
                            }), DispatcherPriority.Background);
                        }
                    },
                    myStreamCts.Token);

                string fullAssistantText;
                lock (tokenAccumulator)
                {
                    fullAssistantText = tokenAccumulator.ToString();
                }

                Dispatcher.Invoke(new Action(() =>
                {
                    streamFinished = true;
                    assistantMsg.Content = fullAssistantText;
                    assistantMsg.IsStreaming = false;

                    // Parse structured spreadsheet actions from Excel responses
                    if (_excelCtrl != null)
                    {
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
                    }
                    else if (_pptCtrl != null)
                    {
                        string cleanContent;
                        var pptActions = PowerPointActionParser.ParseStructuredActions(fullAssistantText, out cleanContent);
                        if (pptActions != null && pptActions.Count > 0)
                        {
                            assistantMsg.Content = cleanContent;
                            foreach (var act in pptActions)
                            {
                                assistantMsg.PowerPointActions.Add(act);
                            }
                            assistantMsg.NotifyPowerPointActionsChanged();
                        }
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
                    "When generating calculations, formulas, analysis, or spreadsheet changes:\n" +
                    "1. Inspect the provided Worksheet Context with its explicit Column Letters (Col A, Col B, Col C, etc.) and Header names.\n" +
                    "2. When matching category/text columns with prefixes (e.g. '0001-non ferrous items'), use wildcard criteria (e.g. \"*non ferrous*\") in SUMIF/COUNTIF.\n" +
                    "3. Start row-level data formulas at Row 2 (e.g. F2, G2) rather than header Row 1.\n" +
                    "4. Propose only bounded A1 cell/range changes (e.g. B2:B100, not unbounded full columns like B:B). Every change is previewed and requires user confirmation.\n" +
                    "5. ALWAYS return executable spreadsheet actions in a structured <excel_actions> XML block when a native change is requested:\n" +
                    "   <excel_actions>\n" +
                    "     <excel_action target=\"K20\" type=\"formula\" formula=\"=SUMIF(B2:B100, &quot;*non ferrous*&quot;, F2:F100)\" description=\"Total non-ferrous value\" />\n" +
                    "     <excel_action target=\"K21\" type=\"formula\" formula=\"=COUNTIF(E2:E100, 0)\" description=\"Count of zero-quantity items\" />\n" +
                    "     <excel_action target=\"K22\" type=\"formula\" formula=\"=AVERAGEIF(E2:E100, 0, F2:F100)\" description=\"Average value of zero-quantity items\" />\n" +
                    "     <excel_action target=\"G2:G27\" type=\"filldown\" formula=\"=IF(F2&gt;50000, &quot;High Value&quot;, &quot;&quot;)\" description=\"High value flag (&gt;50,000)\" />\n" +
                    "   </excel_actions>\n" +
                    "6. Supported action types are formula, value, filldown, table, create_table, conditional_format, sort, filter, data_validation, chart, pivot_table, named_range, and remove_duplicates. " +
                    "Use the value attribute for each action's concise option (for example value=\"descending\", value=\"list:Open,Closed\", or value=\"columns:1,2\") and state any assumptions in the description.\n" +
                    "7. Provide a brief conversational summary above or below the action block without tutorial how-to steps.";
            }
            else if (_hostType == "Word" || _wordCtrl != null)
            {
                hostContext = "\n\nYou are embedded inside Microsoft Word. The provided document context may include prompt-relevant excerpts, a live outline, and action items. " +
                    "When the user asks to write, edit, rewrite, summarize, translate, or review text, provide polished text directly without tutorial meta-commentary. " +
                    "Preserve factual meaning unless the user requests a substantive change. Use a Markdown table when a table is the clearest result. " +
                    "If sources are attached, cite only supplied sources using [Source: filename, page/section]; do not invent citations.";
            }
            else if (_hostType == "PowerPoint" || _pptCtrl != null)
            {
                hostContext = "\n\nYou are embedded inside Microsoft PowerPoint. The supplied context can include the full deck, sections, slide text, and speaker notes. " +
                    "When the user asks for slides or bullet points, provide structured slides with Slide titles, concise bullets, speaker notes, and a one-line Visual suggestion. " +
                    "Preserve the presentation's existing style and structure where possible. For a requested deck reorganization, use only an optional <powerpoint_actions> block with safe action types move_slide, create_section, rename_section, or set_notes; use numbered existing slides and include no other commands.";
            }
            return (basePrompt ?? "You are an expert AI assistant embedded inside Microsoft Office.") + hostContext;
        }

        private PromptContextScope GetPromptContextScope()
        {
            try
            {
                var item = CmbContextScope != null ? CmbContextScope.SelectedItem as ComboBoxItem : null;
                string tag = item != null ? Convert.ToString(item.Tag) : "Selection";
                if (string.Equals(tag, "CurrentFile", StringComparison.OrdinalIgnoreCase)) return PromptContextScope.CurrentFile;
                if (string.Equals(tag, "SelectionAndFile", StringComparison.OrdinalIgnoreCase)) return PromptContextScope.SelectionAndFile;
                if (string.Equals(tag, "AttachmentsOnly", StringComparison.OrdinalIgnoreCase)) return PromptContextScope.AttachmentsOnly;
            }
            catch { }
            return PromptContextScope.Selection;
        }

        private static bool IncludesSelection(PromptContextScope scope)
        {
            return scope == PromptContextScope.Selection || scope == PromptContextScope.SelectionAndFile;
        }

        private static bool IncludesCurrentFile(PromptContextScope scope)
        {
            return scope == PromptContextScope.CurrentFile || scope == PromptContextScope.SelectionAndFile;
        }

        private string ComposePromptWithContext(string prompt, PromptContextScope scope)
        {
            if (string.IsNullOrWhiteSpace(prompt)) return prompt;
            var builder = new StringBuilder(prompt);
            if (IncludesSelection(scope))
            {
                string selected = GetSelectedTextOnly();
                if (!string.IsNullOrWhiteSpace(selected))
                    builder.AppendFormat("\n\n[Selected Context]:\n{0}", selected);
            }
            if (IncludesCurrentFile(scope))
            {
                string currentFileContext = GetCurrentFileContext(prompt);
                if (!string.IsNullOrWhiteSpace(currentFileContext))
                    builder.AppendFormat("\n\n[Current File Context]:\n{0}", currentFileContext);
            }
            return builder.ToString();
        }

        private string GetCurrentFileContext(string prompt)
        {
            try
            {
                if (_wordCtrl != null)
                {
                    string docText = _wordCtrl.GetRelevantDocumentContext(prompt, 24000);
                    if (!string.IsNullOrWhiteSpace(docText))
                    {
                        return string.Format("[Document Content: {0}]:\n{1}", _currentDocumentKey, docText);
                    }
                }
                if (_excelCtrl != null)
                {
                    string snapshot = _excelCtrl.GetWorksheetSnapshot(70, 26);
                    if (!string.IsNullOrWhiteSpace(snapshot))
                    {
                        return snapshot;
                    }
                }
                if (_pptCtrl != null)
                {
                    string reviewContext = _pptCtrl.GetPresentationReviewContext(7000);
                    string deckText = _pptCtrl.GetPresentationText(41000);
                    if (string.IsNullOrWhiteSpace(reviewContext)) return deckText;
                    if (string.IsNullOrWhiteSpace(deckText)) return reviewContext;
                    return reviewContext + "\n\n" + deckText;
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("GetCurrentFileContext failed: {0}", ex.Message));
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

        private void BtnActionHistory_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var entries = ActionAuditStore.Instance.GetRecent(30);
                if (entries == null || entries.Count == 0)
                {
                    MessageBox.Show("No approved AI actions have been recorded on this Windows account yet.",
                        "AI Action Log", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                var log = new StringBuilder();
                log.AppendLine("Recent approved AI actions (stored locally and encrypted):");
                log.AppendLine();
                foreach (var entry in entries)
                {
                    log.AppendFormat("{0:u} | {1} | {2}", entry.TimestampUtc, entry.Host, entry.ActionType);
                    if (!string.IsNullOrWhiteSpace(entry.Target)) log.AppendFormat(" | {0}", entry.Target);
                    log.AppendLine();
                    if (!string.IsNullOrWhiteSpace(entry.Summary)) log.AppendLine(entry.Summary);
                    log.AppendLine();
                }

                MessageBox.Show(log.ToString(), "AI Action Log", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("Could not show action history: {0}", ex.Message));
                MessageBox.Show("The local AI action log could not be opened.", "AI Action Log", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
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

        private void BtnHelp_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                IntPtr ownerHwnd = IntPtr.Zero;
                if (_hostAppObj != null)
                {
                    try { ownerHwnd = OfficeDockedPane.ResolveHostHwnd(_hostAppObj); }
                    catch { }
                }
                HelpWindow.ShowHelp(ownerHwnd);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Could not open the User Manual: {0}", ex.Message), "AI Assistant", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private string GetSelectedModelName()
        {
            try
            {
                string text = CmbModel != null ? (!string.IsNullOrWhiteSpace(CmbModel.Text) ? CmbModel.Text : (CmbModel.SelectedItem as string)) : null;
                return !string.IsNullOrWhiteSpace(text) ? text.Trim() : (ConfigManager.Instance.DefaultModel ?? "default");
            }
            catch
            {
                return ConfigManager.Instance.DefaultModel ?? "default";
            }
        }

        private string GetLastUserPrompt()
        {
            try
            {
                var lastUser = _messages.LastOrDefault(m => m.IsUser);
                return lastUser != null ? lastUser.Content : string.Empty;
            }
            catch
            {
                return string.Empty;
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
                if (!ConfirmSpreadsheetAction(action)) return;
                bool ok = _excelCtrl.ApplySpreadsheetAction(action);
                if (!ok && !string.IsNullOrEmpty(action.ErrorMessage))
                {
                    MessageBox.Show(action.ErrorMessage, "Spreadsheet Action Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else if (ok)
                {
                    ActionAuditStore.Instance.Record(
                        "Excel",
                        action.Type.ToString(),
                        action.Target,
                        DescribeSpreadsheetAction(action),
                        action.IsUndoable,
                        GetLastUserPrompt(),
                        _currentDocumentKey,
                        GetSelectedModelName(),
                        action.Content,
                        !string.IsNullOrEmpty(action.ResultText) ? action.ResultText : "Applied successfully");
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

            var preview = new StringBuilder();
            preview.AppendLine("Review the following Excel changes before applying:");
            bool hasNonUndoable = false;
            foreach (var pendingAction in msg.Actions)
            {
                if (pendingAction.Status == SpreadsheetActionStatus.Pending || pendingAction.Status == SpreadsheetActionStatus.Error)
                {
                    preview.AppendLine(DescribeSpreadsheetAction(pendingAction));
                    if (!pendingAction.IsUndoable) hasNonUndoable = true;
                }
            }
            if (hasNonUndoable)
            {
                preview.AppendLine("\n⚠ Note: One or more actions cannot be undone by Excel Undo.");
            }
            if (MessageBox.Show(preview.ToString(), "Review Spreadsheet Actions", MessageBoxButton.YesNo,
                hasNonUndoable ? MessageBoxImage.Warning : MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            int appliedCount = 0;
            foreach (var act in msg.Actions)
            {
                if (act.Status == SpreadsheetActionStatus.Pending || act.Status == SpreadsheetActionStatus.Error)
                {
                    if (_excelCtrl.ApplySpreadsheetAction(act))
                    {
                        appliedCount++;
                        ActionAuditStore.Instance.Record(
                            "Excel",
                            act.Type.ToString(),
                            act.Target,
                            DescribeSpreadsheetAction(act),
                            act.IsUndoable,
                            GetLastUserPrompt(),
                            _currentDocumentKey,
                            GetSelectedModelName(),
                            act.Content,
                            !string.IsNullOrEmpty(act.ResultText) ? act.ResultText : "Applied successfully");
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

        private void BtnPreviewMessage_Click(object sender, RoutedEventArgs e)
        {
            var msg = GetMessageFromSender(sender);
            if (msg == null || string.IsNullOrWhiteSpace(msg.Content)) return;
            MessageBox.Show(msg.Content, "AI Response Preview", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void BtnInsertMessage_Click(object sender, RoutedEventArgs e)
        {
            var msg = GetMessageFromSender(sender);
            if (msg == null || string.IsNullOrEmpty(msg.Content)) return;
            string content = msg.Content;
            try
            {
                if (!ConfirmInsert(content)) return;
                if (_wordCtrl != null)
                {
                    if (ChkTrackChanges != null && ChkTrackChanges.IsChecked == true)
                    {
                        if (!string.IsNullOrWhiteSpace(_wordCtrl.GetSelectedText()))
                            _wordCtrl.ReplaceSelectionWithTrackChanges(content);
                        else
                            _wordCtrl.InsertTextAtCursorWithTrackChanges(content);
                        ActionAuditStore.Instance.Record(
                            "Word",
                            "Tracked edit",
                            "Selection / cursor",
                            content,
                            true,
                            GetLastUserPrompt(),
                            _currentDocumentKey,
                            GetSelectedModelName(),
                            content,
                            "Inserted with Track Changes");
                    }
                    else
                    {
                        _wordCtrl.InsertTextAtCursor(content);
                        ActionAuditStore.Instance.Record(
                            "Word",
                            "Insert",
                            "Selection / cursor",
                            content,
                            true,
                            GetLastUserPrompt(),
                            _currentDocumentKey,
                            GetSelectedModelName(),
                            content,
                            "Inserted at cursor");
                    }
                }
                else if (_excelCtrl != null)
                {
                    _excelCtrl.InsertText(content);
                    ActionAuditStore.Instance.Record(
                        "Excel",
                        "Insert",
                        "Selection",
                        content,
                        true,
                        GetLastUserPrompt(),
                        _currentDocumentKey,
                        GetSelectedModelName(),
                        content,
                        "Inserted into Excel");
                }
                else if (_pptCtrl != null)
                {
                    _pptCtrl.CreateOrUpdateDeckFromOutline(content);
                    ActionAuditStore.Instance.Record(
                        "PowerPoint",
                        "Create or update slides",
                        "Active deck",
                        content,
                        true,
                        GetLastUserPrompt(),
                        _currentDocumentKey,
                        GetSelectedModelName(),
                        content,
                        "Created or updated slides");
                }
                else
                    MessageBox.Show("No Office document is active. Open a document and try again.", "AI Assistant", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Could not insert text: {0}", ex.Message), "AI Assistant", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool ConfirmSpreadsheetAction(SpreadsheetAction action)
        {
            return MessageBox.Show(DescribeSpreadsheetAction(action), "Review Spreadsheet Action", MessageBoxButton.YesNo,
                action.IsUndoable ? MessageBoxImage.Question : MessageBoxImage.Warning) == MessageBoxResult.Yes;
        }

        private static string DescribeSpreadsheetAction(SpreadsheetAction action)
        {
            if (action == null) return "No spreadsheet action is available.";
            string description = string.IsNullOrWhiteSpace(action.Description) ? "No additional description provided." : action.Description;
            string undoWarning = action.IsUndoable
                ? string.Empty
                : "\n\n⚠ WARNING: This action cannot be reliably undone by Excel Undo.";
            return string.Format("Excel will apply a {0} action to {1}.\n\nDescription: {2}\n\nProposed change:\n{3}{4}",
                action.Type, action.Target, description, action.Content, undoWarning);
        }

        private void BtnApplyPowerPointAction_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            if (btn == null) return;
            var action = btn.Tag as PowerPointAction;
            if (action == null) return;

            if (_pptCtrl != null)
            {
                if (!ConfirmPowerPointAction(action)) return;
                if (_pptCtrl.ApplyPowerPointAction(action))
                {
                    ActionAuditStore.Instance.Record(
                        "PowerPoint",
                        action.Type,
                        action.TargetDisplay,
                        DescribePowerPointAction(action),
                        action.IsUndoable,
                        GetLastUserPrompt(),
                        _currentDocumentKey,
                        GetSelectedModelName(),
                        action.ContentDisplay,
                        !string.IsNullOrEmpty(action.ResultText) ? action.ResultText : "Applied successfully");
                }
            }
            else
            {
                MessageBox.Show("No active PowerPoint presentation found.", "AI Assistant", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void BtnApplyAllPowerPointActions_Click(object sender, RoutedEventArgs e)
        {
            var msg = GetMessageFromSender(sender);
            if (msg == null || msg.PowerPointActions == null || msg.PowerPointActions.Count == 0) return;

            if (_pptCtrl == null)
            {
                MessageBox.Show("No active PowerPoint presentation found.", "AI Assistant", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var preview = new StringBuilder();
            preview.AppendLine("Review the following PowerPoint changes before applying:");
            foreach (var pendingAction in msg.PowerPointActions)
            {
                if (pendingAction.Status == PowerPointActionStatus.Pending || pendingAction.Status == PowerPointActionStatus.Error)
                {
                    preview.AppendLine(DescribePowerPointAction(pendingAction));
                }
            }

            if (MessageBox.Show(preview.ToString(), "Review PowerPoint Actions", MessageBoxButton.YesNo,
                MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            int appliedCount = 0;
            foreach (var act in msg.PowerPointActions)
            {
                if (act.Status == PowerPointActionStatus.Pending || act.Status == PowerPointActionStatus.Error)
                {
                    if (_pptCtrl.ApplyPowerPointAction(act))
                    {
                        appliedCount++;
                        ActionAuditStore.Instance.Record(
                            "PowerPoint",
                            act.Type,
                            act.TargetDisplay,
                            DescribePowerPointAction(act),
                            act.IsUndoable,
                            GetLastUserPrompt(),
                            _currentDocumentKey,
                            GetSelectedModelName(),
                            act.ContentDisplay,
                            !string.IsNullOrEmpty(act.ResultText) ? act.ResultText : "Applied successfully");
                    }
                }
            }
        }

        private bool ConfirmPowerPointAction(PowerPointAction action)
        {
            return MessageBox.Show(DescribePowerPointAction(action), "Review PowerPoint Action", MessageBoxButton.YesNo,
                MessageBoxImage.Question) == MessageBoxResult.Yes;
        }

        private static string DescribePowerPointAction(PowerPointAction action)
        {
            if (action == null) return "No PowerPoint action is available.";
            return string.Format("PowerPoint will apply a {0} action.\n\nDescription: {1}\n\nProposed change:\n{2}",
                action.Type, action.Description, action.ContentDisplay);
        }

        private bool ConfirmInsert(string content)
        {
            string host = _wordCtrl != null ? "Word" : (_excelCtrl != null ? "Excel" : (_pptCtrl != null ? "PowerPoint" : "Office"));
            string operation;
            if (_wordCtrl != null && ChkTrackChanges != null && ChkTrackChanges.IsChecked == true)
                operation = "apply the response as Word Track Changes";
            else if (_pptCtrl != null)
                operation = "create or update slides from the response";
            else if (_excelCtrl != null)
            {
                string cellTarget;
                string cleanContent = ExcelController.ExtractCleanExcelContent(content, out cellTarget);
                string targetLabel = !string.IsNullOrEmpty(cellTarget) ? ("cell " + cellTarget) : "the active selection";

                var tableLines = new List<string>();
                var lines = cleanContent.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var l in lines)
                {
                    string tl = l.Trim();
                    if (tl.StartsWith("|") && tl.EndsWith("|") && !Regex.IsMatch(tl, @"^\|?\s*[-:]+\s*\|[\s-:|]*$"))
                        tableLines.Add(tl);
                }

                if (tableLines.Count > 0)
                {
                    int cols = tableLines[0].Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries).Length;
                    operation = string.Format("write a {0} row × {1} column table into {2}", tableLines.Count, cols, targetLabel);
                }
                else
                {
                    operation = string.Format("write the response into {0}", targetLabel);
                }
            }
            else
                operation = "insert the response";

            string preview = content.Length > 3000 ? content.Substring(0, 3000) + "\n...[preview truncated]" : content;
            return MessageBox.Show(string.Format("{0} will {1}. Review the preview before continuing:\n\n{2}", host, operation, preview),
                "Review AI Change", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes;
        }

        private void BtnUndoChange_Click(object sender, RoutedEventArgs e)
        {
            bool undone = false;
            try
            {
                if (_wordCtrl != null) undone = _wordCtrl.UndoLastChange();
                else if (_excelCtrl != null) undone = _excelCtrl.UndoLastAction();
                else if (_pptCtrl != null) undone = _pptCtrl.UndoLastChange();
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("Undo AI change failed: {0}", ex.Message));
            }

            if (undone)
            {
                ActionAuditStore.Instance.Record(
                    _hostType,
                    "Undo",
                    "Last Office action",
                    "User requested undo of the most recent Office action.",
                    false,
                    "Undo",
                    _currentDocumentKey,
                    GetSelectedModelName(),
                    "Undo last action",
                    "Undone");
            }
            else
            {
                MessageBox.Show("Office could not undo the last change. The native undo stack may be empty or a non-AI edit may be the latest action.",
                    "Undo Unavailable", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void BtnAcceptRevisions_Click(object sender, RoutedEventArgs e)
        {
            ApplyRevisionDecision(true);
        }

        private void BtnRejectRevisions_Click(object sender, RoutedEventArgs e)
        {
            ApplyRevisionDecision(false);
        }

        private void ApplyRevisionDecision(bool accept)
        {
            if (_wordCtrl == null) return;
            int pending = _wordCtrl.GetPendingRevisionCount();
            if (pending <= 0)
            {
                MessageBox.Show("There are no pending tracked revisions in the active Word document.", "No Tracked Edits", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            bool hasSelection = !string.IsNullOrWhiteSpace(_wordCtrl.GetSelectedText());
            string operation = accept ? "accept" : "reject";
            string scope = hasSelection ? "the tracked changes in the selected text" : "all tracked changes in this document";
            if (MessageBox.Show(string.Format("This will {0} {1}.", operation, scope), "Review Tracked Changes",
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            bool applied = false;
            try
            {
                if (hasSelection)
                    applied = accept ? _wordCtrl.AcceptRevisionsInSelection() : _wordCtrl.RejectRevisionsInSelection();
                else
                    applied = accept ? _wordCtrl.AcceptAllRevisions() : _wordCtrl.RejectAllRevisions();
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("Tracked revision decision failed: {0}", ex.Message));
            }

            if (applied)
            {
                ActionAuditStore.Instance.Record(
                    "Word",
                    accept ? "Accept tracked edits" : "Reject tracked edits",
                    hasSelection ? "Selection" : "Document",
                    string.Format("User chose to {0} tracked revisions.", operation),
                    true,
                    operation + " revisions",
                    _currentDocumentKey,
                    GetSelectedModelName(),
                    operation,
                    applied ? "Success" : "Failed");
            }
            else
            {
                MessageBox.Show("Word could not apply the tracked revision decision. Try selecting the revision or using Word's Review tab.",
                    "Tracked Changes", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private async void BtnQuickAction_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            if (btn != null && btn.Tag is string)
            {
                string prompt = (string)btn.Tag;
                string title = btn.Content != null ? btn.Content.ToString() : null;
                PromptContextScope scope = GetPromptContextScope();
                if (_pptCtrl != null && (btn.Name == "BtnQuickReview" || btn.Name == "BtnQuickDeck"))
                    scope = PromptContextScope.SelectionAndFile;
                await ExecuteExternalPromptAsync(prompt, title, scope);
            }
        }

        private void BtnInsertVisual_Click(object sender, RoutedEventArgs e)
        {
            if (_pptCtrl == null)
            {
                MessageBox.Show("Open PowerPoint and select a slide before inserting a visual.", "Insert Image", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Insert Image into Active Slide",
                Multiselect = false,
                Filter = "Image files (*.png;*.jpg;*.jpeg;*.webp;*.bmp)|*.png;*.jpg;*.jpeg;*.webp;*.bmp|All files (*.*)|*.*"
            };
            if (dialog.ShowDialog() != true || string.IsNullOrWhiteSpace(dialog.FileName)) return;

            string altText = Path.GetFileNameWithoutExtension(dialog.FileName);
            if (MessageBox.Show(string.Format("Insert '{0}' into the active slide?\n\nAlt text: {1}", Path.GetFileName(dialog.FileName), altText),
                "Review Image Insertion", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            if (_pptCtrl.InsertImageFromFile(dialog.FileName, altText))
            {
                ActionAuditStore.Instance.Record(
                    "PowerPoint",
                    "Insert image",
                    "Active slide",
                    Path.GetFileName(dialog.FileName),
                    true,
                    "Insert image",
                    _currentDocumentKey,
                    GetSelectedModelName(),
                    dialog.FileName,
                    "Image inserted");
            }
            else
            {
                MessageBox.Show("PowerPoint could not insert the selected image.", "Insert Image", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void TxtInput_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
            {
                e.Handled = true;
                BtnSend_Click(sender, e);
            }
        }

        private void TxtInput_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            // Excel 2010 can retain native focus after clicking a child WPF control.  Set focus
            // synchronously so the very first keystroke is already targeted at the WPF input source.
            IntPtr promptWindow = GetPromptInputWindowHandle();
            if (promptWindow != IntPtr.Zero)
                NativeWnd.SetFocus(promptWindow);

            FocusPromptInput();
            EventHandler handler = PromptInputFocusRequested;
            if (handler != null) handler(this, EventArgs.Empty);
        }

        private void TxtInput_GotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            // Arm routing on every focus path (Tab, programmatic focus), not just mouse-down.
            FocusPromptInput();
            EventHandler handler = PromptInputFocusRequested;
            if (handler != null) handler(this, EventArgs.Empty);
        }

        private void TxtInput_LostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            // Only disarm when focus leaves the whole sidebar; moving to another pane control
            // (e.g. the model combo) must keep routing active.
            if (IsKeyboardFocusWithin) return;
            EventHandler handler = PromptInputFocusLost;
            if (handler != null) handler(this, EventArgs.Empty);
        }

        private void ChatSidebar_Loaded(object sender, RoutedEventArgs e)
        {
            // Attach runs before the WPF visual is connected; re-arm once Loaded fires so the
            // pane can cache the now-valid HwndSource.
            EventHandler handler = PromptInputFocusRequested;
            if (handler != null) handler(this, EventArgs.Empty);
        }

        public bool IsPromptKeyboardFocused
        {
            get { return TxtInput != null && (TxtInput.IsKeyboardFocused || TxtInput.IsKeyboardFocusWithin || IsKeyboardFocusWithin); }
        }

        private IntPtr _cachedPromptHwnd = IntPtr.Zero;

        public IntPtr GetPromptInputWindowHandle()
        {
            if (TxtInput == null) return IntPtr.Zero;
            if (_cachedPromptHwnd != IntPtr.Zero && NativeWnd.IsWindow(_cachedPromptHwnd))
                return _cachedPromptHwnd;
            _cachedPromptHwnd = IntPtr.Zero;
            HwndSource source = PresentationSource.FromVisual(TxtInput) as HwndSource;
            if (source != null) _cachedPromptHwnd = source.Handle;
            return _cachedPromptHwnd;
        }

        public void FocusPromptInput()
        {
            try
            {
                if (TxtInput == null) return;
                Dispatcher.BeginInvoke(new Action(delegate
                {
                    IntPtr promptWindow = GetPromptInputWindowHandle();
                    if (promptWindow != IntPtr.Zero)
                        NativeWnd.SetFocus(promptWindow);

                    TxtInput.Focus();
                    Keyboard.Focus(TxtInput);
                }), DispatcherPriority.Input);
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("ChatSidebar.FocusPromptInput failed: {0}", ex.Message));
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
