using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using MistralOfficeAddin.API;
using MistralOfficeAddin.API.Models;
using MistralOfficeAddin.Core;
using MistralOfficeAddin.Hosts;

namespace MistralOfficeAddin.UI
{
    public partial class ChatSidebar : UserControl
    {
        private readonly ObservableCollection<ChatMessage> _messages = new ObservableCollection<ChatMessage>();
        private CancellationTokenSource _streamingCts;
        private string _currentDocumentKey = "OfficeSession";
        private object _hostAppObj;
        private string _hostType = "Office";

        // Host Controllers — created lazily, not during startup
        private WordController _wordCtrl;
        private ExcelController _excelCtrl;
        private PowerPointController _pptCtrl;
        private OutlookController _outlookCtrl;
        private bool _hostInitialized;

        public ChatSidebar()
        {
            InitializeComponent();
            MessagesItemsControl.ItemsSource = _messages;

            SelectConfiguredModel();

            // Show welcome message immediately without waiting for host init
            _messages.Add(new ChatMessage("system", "Mistral AI Assistant is ready. Click Configure (⚙️) to enter your API key, then start chatting!"));
        }

        private void SelectConfiguredModel()
        {
            try
            {
                string configured = ConfigManager.Instance.DefaultModel;
                if (string.IsNullOrWhiteSpace(configured) || CmbModel == null) return;
                foreach (ComboBoxItem item in CmbModel.Items)
                {
                    if (string.Equals(Convert.ToString(item.Content), configured, StringComparison.OrdinalIgnoreCase))
                    {
                        CmbModel.SelectedItem = item;
                        return;
                    }
                }
            }
            catch { }
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
                else if (string.Equals(_hostType, "Outlook", StringComparison.OrdinalIgnoreCase))
                {
                    _outlookCtrl = new OutlookController(_hostAppObj);
                    docName = _outlookCtrl.GetActiveItemTitle();
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
                    _messages.Add(new ChatMessage("system", string.Format("Welcome to Mistral AI for {0}! Ask anything or use ribbon buttons to draft, rewrite, or summarize.", _hostType)));
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

            string contextText = GetCurrentContextText(false);
            string fullPrompt = string.IsNullOrEmpty(contextText)
                ? prompt
                : string.Format("{0}\n\n[Context]:\n{1}", prompt, contextText);

            await SendMessageAsync(fullPrompt, promptTitle);
        }

        private async void BtnSend_Click(object sender, RoutedEventArgs e)
        {
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
            if (string.IsNullOrWhiteSpace(config.ApiKey))
            {
                _messages.Add(new ChatMessage("system", "API Key is missing. Click ⚙️ Settings to configure your Mistral API key."));
                ScrollToBottom();
                return;
            }

            string displayUserMessage = customDisplayTitle ?? promptToSend;
            _messages.Add(new ChatMessage("user", displayUserMessage));
            ScrollToBottom();

            var assistantMsg = new ChatMessage("assistant", "") { IsStreaming = true };
            _messages.Add(assistantMsg);
            ScrollToBottom();

            BtnSend.IsEnabled = false;
            TypingIndicator.Visibility = Visibility.Visible;

            if (_streamingCts != null)
                _streamingCts.Cancel();
            _streamingCts = new CancellationTokenSource();

            string selectedModel = config.DefaultModel ?? "mistral-large-latest";
            if (CmbModel.SelectedItem is ComboBoxItem)
            {
                var item = (ComboBoxItem)CmbModel.SelectedItem;
                selectedModel = Convert.ToString(item.Content);
            }

            try
            {
                var historyForApi = _messages
                    .Where(m => (m.IsUser || m.IsAssistant) && m != assistantMsg)
                    .ToList();

                var boundedMessages = TokenCounter.TruncateToFit(historyForApi, 24000, config.SystemPrompt);

                using (var client = new MistralClient(config.BaseUrl, config.ApiKey))
                {
                    await client.StreamChatCallbackAsync(
                        selectedModel,
                        boundedMessages,
                        config.Temperature,
                        config.MaxTokens,
                        delta =>
                        {
                            // Use Invoke with Background priority — prevents UI thread lockup
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                assistantMsg.Content += delta;
                            }), DispatcherPriority.Background);
                        },
                        _streamingCts.Token);
                }

                assistantMsg.IsStreaming = false;
                // Scroll once at the end, not on every token
                ScrollToBottom();
                SaveConversationHistory();
            }
            catch (OperationCanceledException)
            {
                assistantMsg.Content += "\n\n*(Generation stopped)*";
                assistantMsg.IsStreaming = false;
            }
            catch (Exception ex)
            {
                Logger.Error("Chat completion error", ex);
                assistantMsg.Content = string.Format("Error: {0}", ex.Message);
                assistantMsg.IsStreaming = false;
            }
            finally
            {
                BtnSend.IsEnabled = true;
                TypingIndicator.Visibility = Visibility.Collapsed;
                ScrollToBottom();
            }
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
                if (selectionOnly && _excelCtrl == null && _pptCtrl == null && _outlookCtrl == null)
                {
                    return string.Empty;
                }
                if (_excelCtrl != null) return _excelCtrl.GetSelectedRangeValues();
                if (_pptCtrl != null) return _pptCtrl.GetSlideText();
                if (_outlookCtrl != null) return _outlookCtrl.GetEmailBody();
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
                _streamingCts.Cancel();
        }

        public void StartNewChat(bool confirm)
        {
            if (confirm)
            {
                if (MessageBox.Show("Start a new chat session and clear current conversation?", "Mistral AI",
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
                win.ShowDialog();
            }
            catch (Exception ex)
            {
                Logger.Error("Failed to open SettingsWindow", ex);
            }
        }

        private ChatMessage GetMessageFromSender(object sender)
        {
            var btn = sender as Button;
            if (btn == null) return null;
            return btn.DataContext as ChatMessage;
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
                    _pptCtrl.AddBulletPoints(content.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).ToList());
                else if (_outlookCtrl != null)
                    _outlookCtrl.SetComposeBody(content);
                else
                    MessageBox.Show("No Office document is active. Open a document and try again.", "Mistral AI", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Could not insert text: {0}", ex.Message), "Mistral AI", MessageBoxButton.OK, MessageBoxImage.Error);
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
            if (CmbModel != null && CmbModel.SelectedItem is ComboBoxItem)
            {
                var item = (ComboBoxItem)CmbModel.SelectedItem;
                ConfigManager.Instance.DefaultModel = Convert.ToString(item.Content);
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
