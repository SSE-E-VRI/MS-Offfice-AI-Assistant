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
using MSOfficeAIAssistant.Core.Actions;
using MSOfficeAIAssistant.Core.Session;
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

        private readonly ObservableCollection<ChatMessage> _messages = new ObservableCollection<ChatMessage>();
        private readonly ObservableCollection<AttachmentItemViewModel> _pendingAttachments = new ObservableCollection<AttachmentItemViewModel>();
        private readonly ChatOrchestrator _orchestrator;
        private readonly AssistantSession _session;
        private string _currentDocumentKey = "OfficeSession";
        private object _hostAppObj;
        private string _hostType = "Office";

        // Host Controllers — created lazily, not during startup
        private IOfficeHostController _hostController;
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
            _session = new AssistantSession(_orchestrator, _messages);
            _session.HostType = _hostType;

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
                if (_session != null && _session.Orchestrator != null)
                {
                    var newProvider = ProviderFactory.CreateFromConfig(ConfigManager.Instance);
                    _session.Orchestrator.UpdateProvider(newProvider);
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
                    _hostController = _wordCtrl;
                }
                else if (string.Equals(_hostType, "Excel", StringComparison.OrdinalIgnoreCase))
                {
                    _excelCtrl = new ExcelController(_hostAppObj);
                    _hostController = _excelCtrl;
                }
                else if (string.Equals(_hostType, "PowerPoint", StringComparison.OrdinalIgnoreCase))
                {
                    _pptCtrl = new PowerPointController(_hostAppObj);
                    _hostController = _pptCtrl;
                }

                if (_hostController != null)
                {
                    docName = _hostController.GetActiveDocumentName();
                }

                _currentDocumentKey = string.IsNullOrWhiteSpace(docName) ? "Document" : docName;
                _session.HostType = _hostType;
                _session.CurrentDocumentKey = _currentDocumentKey;
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
            _session.LoadHistory(_currentDocumentKey);
            if (_messages.Count == 0)
            {
                _messages.Add(new ChatMessage("system", string.Format("Welcome to AI Assistant for {0}! Ask anything or use ribbon buttons to draft, rewrite, or summarize.", _hostType)));
            }
            ScrollToBottom();
        }

        private void SaveConversationHistory()
        {
            _session.SaveHistory();
        }

        public async void ExecuteExternalPrompt(string prompt, string promptTitle)
        {
            await ExecuteExternalPromptAsync(prompt, promptTitle, GetPromptContextScope());
        }

        private async Task ExecuteExternalPromptAsync(string prompt, string promptTitle, PromptContextScope scope)
        {
            if (string.IsNullOrWhiteSpace(prompt)) return;

            var coordinator = _session.StreamCoordinator;
            if (coordinator.IsSending)
            {
                coordinator.Cancel();
                await Task.Delay(100).ConfigureAwait(true);
            }
            if (coordinator.IsSending) return;

            string selectedText = PromptAssembler.IncludesSelection(scope) ? GetSelectedTextOnly() : string.Empty;
            string documentContext = PromptAssembler.IncludesCurrentFile(scope) ? GetCurrentFileContext(prompt) : string.Empty;
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
                if (_hostController != null)
                {
                    string sel = _hostController.GetSelectedText();
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
            if (_session.StreamCoordinator.IsSending) return;
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

            var coordinator = _session.StreamCoordinator;
            if (coordinator.IsSending)
            {
                coordinator.Cancel();
            }

            string displayUserMessage = customDisplayTitle ?? promptToSend;
            var userMsg = new ChatMessage("user", displayUserMessage) { FullContent = promptToSend };
            _messages.Add(userMsg);
            ScrollToBottom();

            var assistantMsg = new ChatMessage("assistant", "") { IsStreaming = true };
            _messages.Add(assistantMsg);
            ScrollToBottom();

            BtnSend.IsEnabled = false;
            TypingIndicator.Visibility = Visibility.Visible;

            var streamCts = coordinator.BeginStream();

            string selectedModel = !string.IsNullOrWhiteSpace(CmbModel.Text)
                ? CmbModel.Text.Trim()
                : (config.DefaultModel ?? "default");

            try
            {
                var attachmentPaths = _pendingAttachments.Select(a => a.FilePath).ToList();
                var prepared = await _session.PreparePayloadAsync(selectedModel, attachmentPaths, assistantMsg);

                if (prepared.DroppedImagesCount > 0)
                {
                    _messages.Insert(_messages.IndexOf(assistantMsg), new ChatMessage("system", "⚠ Note: Image attachments were omitted — the selected model does not support vision analysis."));
                }

                await _session.Orchestrator.StreamChatAsync(
                    prepared.Request,
                    delta =>
                    {
                        coordinator.AccumulateDelta(delta, snapshot =>
                        {
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                if (!coordinator.IsStreamFinished)
                                {
                                    assistantMsg.Content = snapshot;
                                }
                            }), DispatcherPriority.Background);
                        });
                    },
                    streamCts.Token);

                string fullAssistantText = coordinator.FinishStream();

                Dispatcher.Invoke(new Action(() =>
                {
                    _session.ProcessAssistantResponse(fullAssistantText, assistantMsg);

                    ScrollToBottom();
                    _session.SaveHistory();

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
                coordinator.EndSending();
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

        private string ComposePromptWithContext(string prompt, PromptContextScope scope)
        {
            string selected = PromptAssembler.IncludesSelection(scope) ? GetSelectedTextOnly() : null;
            string currentFileContext = PromptAssembler.IncludesCurrentFile(scope) ? GetCurrentFileContext(prompt) : null;
            return PromptAssembler.ComposePromptWithContext(prompt, scope, selected, currentFileContext);
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
            if (_session != null)
            {
                _session.StreamCoordinator.Cancel();
                if (_session.Orchestrator != null)
                {
                    _session.Orchestrator.CancelCurrentStream();
                }
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

        private void BtnApplyOfficeAction_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            var action = btn != null ? btn.Tag as OfficeAction : null;
            if (action == null) return;

            if (!ConfirmOfficeAction(action)) return;
            ExecuteOfficeAction(action);
        }

        private void BtnApplyAllOfficeActions_Click(object sender, RoutedEventArgs e)
        {
            var msg = GetMessageFromSender(sender);
            if (msg == null || msg.OfficeActions == null || msg.OfficeActions.Count == 0) return;

            var pendingActions = msg.OfficeActions.Where(a => a.Status == OfficeActionStatus.Pending || a.Status == OfficeActionStatus.Failed).ToList();
            if (pendingActions.Count == 0) return;

            var preview = new StringBuilder();
            preview.AppendLine("Review the following Office actions before applying:\n");
            bool hasNonUndoable = false;
            foreach (var act in pendingActions)
            {
                preview.AppendLine(DescribeOfficeAction(act));
                preview.AppendLine();
                if (!act.IsUndoable) hasNonUndoable = true;
            }

            if (MessageBox.Show(preview.ToString().TrimEnd(), "Review All Actions", MessageBoxButton.YesNo,
                hasNonUndoable ? MessageBoxImage.Warning : MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            foreach (var act in pendingActions)
            {
                ExecuteOfficeAction(act);
            }
        }

        private bool ExecuteOfficeAction(OfficeAction action)
        {
            if (action == null) return false;

            string host = !string.IsNullOrEmpty(action.Host) ? action.Host : _hostType;

            if (string.Equals(host, "Excel", StringComparison.OrdinalIgnoreCase) || (action.Operation != null && action.Operation.StartsWith("excel.", StringComparison.OrdinalIgnoreCase)))
            {
                if (_excelCtrl == null)
                {
                    MessageBox.Show("No active Excel workbook found.", "AI Assistant", MessageBoxButton.OK, MessageBoxImage.Information);
                    return false;
                }

                var sa = action.ToSpreadsheetAction();
                if (sa == null)
                {
                    MessageBox.Show(string.Format("Unsupported spreadsheet action '{0}'.", action.Operation), "Spreadsheet Action Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }

                var res = _excelCtrl.ExecuteSpreadsheetAction(sa);
                if (res.Success)
                {
                    string msgText = res.Value != null ? Convert.ToString(res.Value) : "Applied successfully";
                    action.Status = OfficeActionStatus.Applied;
                    action.ResultText = msgText;
                    action.ErrorMessage = null;
                    RecordOfficeActionAudit(action, msgText);
                    return true;
                }
                else
                {
                    action.Status = OfficeActionStatus.Failed;
                    action.ErrorMessage = res.ErrorMessage;
                    MessageBox.Show(res.ErrorMessage ?? "Spreadsheet action failed.", "Spreadsheet Action Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }
            }
            else if (string.Equals(host, "PowerPoint", StringComparison.OrdinalIgnoreCase) || (action.Operation != null && action.Operation.StartsWith("powerpoint.", StringComparison.OrdinalIgnoreCase)))
            {
                if (_pptCtrl == null)
                {
                    MessageBox.Show("No active PowerPoint presentation found.", "AI Assistant", MessageBoxButton.OK, MessageBoxImage.Information);
                    return false;
                }

                var pa = action.ToPowerPointAction();
                if (pa == null)
                {
                    MessageBox.Show(string.Format("Unsupported PowerPoint action '{0}'.", action.Operation), "PowerPoint Action Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }

                var res = _pptCtrl.ExecutePowerPointAction(pa);
                if (res.Success)
                {
                    string msgText = res.Value != null ? Convert.ToString(res.Value) : "Applied successfully";
                    action.Status = OfficeActionStatus.Applied;
                    action.ResultText = msgText;
                    action.ErrorMessage = null;
                    RecordOfficeActionAudit(action, msgText);
                    return true;
                }
                else
                {
                    action.Status = OfficeActionStatus.Failed;
                    action.ErrorMessage = res.ErrorMessage;
                    MessageBox.Show(res.ErrorMessage ?? "PowerPoint action failed.", "PowerPoint Action Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }
            }
            else if (string.Equals(host, "Word", StringComparison.OrdinalIgnoreCase) || (action.Operation != null && action.Operation.StartsWith("word.", StringComparison.OrdinalIgnoreCase)))
            {
                if (_wordCtrl == null)
                {
                    MessageBox.Show("No active Word document found.", "AI Assistant", MessageBoxButton.OK, MessageBoxImage.Information);
                    return false;
                }

                HostOperationResult res = null;
                if (string.Equals(action.Operation, "word.add_comment", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(action.Operation, "add_comment", StringComparison.OrdinalIgnoreCase))
                {
                    string comment = action.Parameters != null && action.Parameters.ContainsKey("comment_text") ? Convert.ToString(action.Parameters["comment_text"]) : "";
                    string targetText = action.Parameters != null && action.Parameters.ContainsKey("target_text") ? Convert.ToString(action.Parameters["target_text"]) : null;
                    res = _wordCtrl.ExecuteAddComment(comment, targetText);
                }
                else if (string.Equals(action.Operation, "word.insert_table", StringComparison.OrdinalIgnoreCase) ||
                         string.Equals(action.Operation, "insert_table", StringComparison.OrdinalIgnoreCase))
                {
                    int rows = action.Parameters != null && action.Parameters.ContainsKey("rows") ? Convert.ToInt32(action.Parameters["rows"]) : 2;
                    int cols = action.Parameters != null && action.Parameters.ContainsKey("cols") ? Convert.ToInt32(action.Parameters["cols"]) : 2;
                    res = _wordCtrl.ExecuteInsertTable(rows, cols);
                }
                else
                {
                    MessageBox.Show(string.Format("Unsupported Word action '{0}'.", action.Operation), "Word Action Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }

                if (res != null)
                {
                    if (res.Success)
                    {
                        string msgText = res.Value != null ? Convert.ToString(res.Value) : "Applied successfully";
                        action.Status = OfficeActionStatus.Applied;
                        action.ResultText = msgText;
                        action.ErrorMessage = null;
                        RecordOfficeActionAudit(action, msgText);
                        return true;
                    }
                    else
                    {
                        action.Status = OfficeActionStatus.Failed;
                        action.ErrorMessage = res.ErrorMessage;
                        MessageBox.Show(res.ErrorMessage ?? "Word action failed.", "Word Action Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return false;
                    }
                }
            }

            MessageBox.Show(string.Format("Unsupported action '{0}' for host {1}.", action.Operation, host), "AI Assistant", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        private bool ConfirmOfficeAction(OfficeAction action)
        {
            if (action == null) return false;
            return MessageBox.Show(
                DescribeOfficeAction(action),
                string.Format("Review {0} Action", action.Host ?? _hostType),
                MessageBoxButton.YesNo,
                action.IsUndoable ? MessageBoxImage.Question : MessageBoxImage.Warning) == MessageBoxResult.Yes;
        }

        private string DescribeOfficeAction(OfficeAction action)
        {
            if (action == null) return "No Office action is available.";
            string host = !string.IsNullOrEmpty(action.Host) ? action.Host : _hostType;
            string desc = !string.IsNullOrWhiteSpace(action.PreviewDescription) ? action.PreviewDescription : "No description provided.";
            string undoWarning = action.IsUndoable
                ? string.Empty
                : "\n\n⚠ WARNING: This action cannot be reliably undone by Office Undo.";
            return string.Format("{0} will apply a {1} action on {2}.\n\nDescription: {3}\n\nProposed change:\n{4}{5}",
                host,
                action.Operation,
                action.TargetDisplay,
                desc,
                action.ContentDisplay,
                undoWarning);
        }

        private void RecordOfficeActionAudit(OfficeAction action, string resultText)
        {
            ActionAuditStore.Instance.Record(
                action.Host ?? _hostType,
                action.Operation ?? "Action",
                action.TargetDisplay,
                DescribeOfficeAction(action),
                action.IsUndoable,
                GetLastUserPrompt(),
                _currentDocumentKey,
                GetSelectedModelName(),
                action.ContentDisplay,
                !string.IsNullOrEmpty(resultText) ? resultText : "Applied successfully");
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

        private bool ConfirmInsert(string content)
        {
            string host = _hostController != null ? _hostController.HostType : "Office";
            string operation;
            if (_wordCtrl != null && ChkTrackChanges != null && ChkTrackChanges.IsChecked == true)
                operation = "apply the response as Word Track Changes";
            else if (_pptCtrl != null)
            {
                var slides = PowerPointActionParser.ParseSlideData(content);
                if (slides.Count > 0)
                {
                    var titles = new List<string>();
                    for (int i = 0; i < slides.Count; i++)
                    {
                        string t = !string.IsNullOrWhiteSpace(slides[i].Title) ? slides[i].Title : string.Format("Slide {0}", i + 1);
                        titles.Add(string.Format("{0}. {1}", i + 1, t));
                    }
                    operation = string.Format("create {0} slides:\n{1}", slides.Count, string.Join("\n", titles.ToArray()));
                }
                else
                {
                    operation = "create or update slides from the response";
                }
            }
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
                if (_hostController != null)
                {
                    undone = _hostController.Undo();
                }
                else if (_wordCtrl != null) undone = _wordCtrl.UndoLastChange();
                else if (_excelCtrl != null) undone = _excelCtrl.UndoLastAction();
                else if (_pptCtrl != null) undone = _pptCtrl.Undo();
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
                    "Undo requested",
                    "Undone successfully");
            }
            else
            {
                MessageBox.Show(string.Format("Could not undo the last change in {0}.", _hostType), "AI Assistant", MessageBoxButton.OK, MessageBoxImage.Information);
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
                if (_pptCtrl != null && btn.Name == "BtnQuickDeck")
                {
                    prompt = PromptAssembler.BuildBriefingDeckPrompt(null, null, 5);
                    scope = PromptContextScope.SelectionAndFile;
                }
                else if (_pptCtrl != null && btn.Name == "BtnQuickReview")
                {
                    scope = PromptContextScope.SelectionAndFile;
                }
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
