using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

namespace MistralOfficeAddin
{
    /// <summary>
    /// Chat task pane hosted by Word/Excel/PowerPoint custom task panes.
    /// Registered as an ActiveX control (MistralAI.ChatPane ProgID) via
    /// ComRegisterFunction which writes the Implemented Categories keys
    /// that Office requires for CTP hosting.
    ///
    /// Implements IObjectSafety so Office accepts the control without
    /// security prompts.
    ///
    /// All UI events run on the host's main STA thread, so Office object
    /// model calls are safe directly in handlers; only HTTP runs in
    /// the background.
    /// </summary>
    [Guid("9C4E7A15-2D6B-4F83-B5C9-7A2E1D4F6B83")]
    [ProgId("MistralAI.ChatPane")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class TaskPaneControl : UserControl, IObjectSafety
    {
        // IObjectSafety flags
        private const int INTERFACESAFE_FOR_UNTRUSTED_CALLER = 0x00000001;
        private const int INTERFACESAFE_FOR_UNTRUSTED_DATA   = 0x00000002;
        private const int S_OK = 0;

        private object _appObj;
        private string _host = "Office";
        private SynchronizationContext _syncContext;

        private readonly List<ChatMessage> _history = new List<ChatMessage>();
        private const int MaxHistoryMessages = 30;
        private const int MaxContextChars = 12000;
        private string _lastResponse;
        private bool _busy;

        private Label _lblTitle;
        private LinkLabel _lnkSettings;
        private CheckBox _chkContext;
        private Panel _chatPanel;
        private TextBox _txtInput;
        private Button _btnSend;
        private Button _btnInsert;
        private Button _btnClear;
        private Label _lblStatus;

        public TaskPaneControl()
        {
            BuildUi();
        }

        public void Initialize(object appObj, string host)
        {
            _appObj = appObj;
            _host = host;
            // Capture the UI SynchronizationContext for safe cross-thread marshaling
            _syncContext = SynchronizationContext.Current;
            _lblTitle.Text = "Mistral AI \u2014 " + host;
            Settings st = Settings.Load();
            _chkContext.Checked = st.IncludeContextByDefault;
            if (_history.Count == 0)
            {
                AddBubble("Mistral", "Hi! Ask me anything about your " + DocNoun() +
                    ".\r\nTip: tick \"Include document context\" so I can see what you are working on, " +
                    "then use Insert to place my answer into the " + DocNoun() + ".", false);
            }
        }

        #region IObjectSafety — allow Office to host us without security prompts

        public void GetInterfaceSafetyOptions(ref Guid riid,
            out int pdwSupportedOptions, out int pdwEnabledOptions)
        {
            pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER | INTERFACESAFE_FOR_UNTRUSTED_DATA;
            pdwEnabledOptions   = INTERFACESAFE_FOR_UNTRUSTED_CALLER | INTERFACESAFE_FOR_UNTRUSTED_DATA;
        }

        public void SetInterfaceSafetyOptions(ref Guid riid,
            int dwOptionSetMask, int dwEnabledOptions)
        {
            // Accept whatever the host requests.
        }

        #endregion

        #region ActiveX Control Registration / Unregistration

        // These static methods are called by RegAsm to write the ActiveX
        // "Implemented Categories" registry keys that Office CTP factory
        // requires before it will instantiate our control.

        private static readonly string ControlGuid = "{9C4E7A15-2D6B-4F83-B5C9-7A2E1D4F6B83}";

        [ComRegisterFunction]
        public static void RegisterControl(Type t)
        {
            if (t == null || t.GUID.ToString("B").ToUpperInvariant() != ControlGuid) return;
            try
            {
                string keyPath = "CLSID\\" + ControlGuid;
                using (RegistryKey key = Registry.ClassesRoot.OpenSubKey(keyPath, true))
                {
                    if (key == null) return;

                    // Mark as insertable object
                    using (RegistryKey cats = key.CreateSubKey("Implemented Categories"))
                    {
                        // CATID_SafeForScripting
                        cats.CreateSubKey("{7DD95801-9882-11CF-9FA9-00AA006C42C4}").Close();
                        // CATID_SafeForInitializing
                        cats.CreateSubKey("{7DD95802-9882-11CF-9FA9-00AA006C42C4}").Close();
                        // CATID_InsertableObject (allows CreateCTP to find us)
                        cats.CreateSubKey("{40FC6ED4-2438-11CF-A3DB-080036F12502}").Close();
                    }

                    // "Control" key marks this as a UI control
                    key.CreateSubKey("Control").Close();

                    // "MiscStatus" — default activation flags
                    using (RegistryKey ms = key.CreateSubKey("MiscStatus"))
                    {
                        ms.SetValue("", "131473"); // OLEMISC_RECOMPOSEONRESIZE | etc.
                    }
                }
            }
            catch
            {
                // Fail silently — HKCU registration may not have HKCR access;
                // register.cmd handles the HKCU path separately.
            }
        }

        [ComUnregisterFunction]
        public static void UnregisterControl(Type t)
        {
            if (t == null || t.GUID.ToString("B").ToUpperInvariant() != ControlGuid) return;
            try
            {
                string keyPath = "CLSID\\" + ControlGuid;
                using (RegistryKey key = Registry.ClassesRoot.OpenSubKey(keyPath, true))
                {
                    if (key == null) return;
                    try { key.DeleteSubKeyTree("Implemented Categories"); } catch { }
                    try { key.DeleteSubKey("Control"); } catch { }
                    try { key.DeleteSubKey("MiscStatus"); } catch { }
                }
            }
            catch { }
        }

        #endregion

        private string DocNoun()
        {
            if (_host == "Word") return "document";
            if (_host == "Excel") return "workbook";
            if (_host == "PowerPoint") return "presentation";
            return "file";
        }

        #region UI construction (no designer file)

        private void BuildUi()
        {
            // Header (top)
            Panel header = new Panel();
            header.Dock = DockStyle.Top;
            header.Height = 64;
            header.Padding = new Padding(8, 6, 8, 2);

            _lblTitle = new Label();
            _lblTitle.Text = "Mistral AI";
            _lblTitle.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            _lblTitle.AutoSize = true;
            _lblTitle.Location = new Point(10, 6);
            header.Controls.Add(_lblTitle);

            _lnkSettings = new LinkLabel();
            _lnkSettings.Text = "Settings";
            _lnkSettings.AutoSize = true;
            _lnkSettings.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _lnkSettings.Location = new Point(300, 7);
            _lnkSettings.LinkClicked += delegate { ShowSettings(); };
            header.Controls.Add(_lnkSettings);

            _chkContext = new CheckBox();
            _chkContext.Text = "Include document context";
            _chkContext.AutoSize = true;
            _chkContext.Location = new Point(10, 30);
            _chkContext.Checked = true;
            header.Controls.Add(_chkContext);

            // Chat area (fill)
            _chatPanel = new Panel();
            _chatPanel.Dock = DockStyle.Fill;
            _chatPanel.AutoScroll = true;
            _chatPanel.BackColor = Color.White;
            _chatPanel.Padding = new Padding(6);

            // Bottom (input + buttons)
            Panel bottom = new Panel();
            bottom.Dock = DockStyle.Bottom;
            bottom.Height = 112;

            _txtInput = new TextBox();
            _txtInput.Multiline = true;
            _txtInput.ScrollBars = ScrollBars.Vertical;
            _txtInput.Dock = DockStyle.Top;
            _txtInput.Height = 68;
            _txtInput.Font = new Font("Segoe UI", 9F);
            _txtInput.KeyDown += TxtInputKeyDown;
            bottom.Controls.Add(_txtInput);

            Panel row = new Panel();
            row.Dock = DockStyle.Fill;
            row.Padding = new Padding(0, 6, 0, 0);

            _btnSend = new Button();
            _btnSend.Text = "Send";
            _btnSend.Size = new Size(80, 30);
            _btnSend.Location = new Point(0, 8);
            _btnSend.Click += delegate { Send(); };
            row.Controls.Add(_btnSend);

            _btnInsert = new Button();
            _btnInsert.Text = "Insert";
            _btnInsert.Size = new Size(80, 30);
            _btnInsert.Location = new Point(88, 8);
            _btnInsert.Click += delegate { InsertLastResponse(); };
            row.Controls.Add(_btnInsert);

            _btnClear = new Button();
            _btnClear.Text = "Clear";
            _btnClear.Size = new Size(70, 30);
            _btnClear.Location = new Point(176, 8);
            _btnClear.Click += delegate { ClearChat(); };
            row.Controls.Add(_btnClear);

            _lblStatus = new Label();
            _lblStatus.Text = "Ready.";
            _lblStatus.AutoSize = true;
            _lblStatus.ForeColor = SystemColors.GrayText;
            _lblStatus.Location = new Point(254, 15);
            _lblStatus.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            row.Controls.Add(_lblStatus);

            bottom.Controls.Add(row);

            // Docking order matters: fill control first, then sides.
            Controls.Add(_chatPanel);
            Controls.Add(bottom);
            Controls.Add(header);
            BackColor = SystemColors.Control;
        }

        #endregion

        #region Chat flow

        private void TxtInputKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && (e.Control || e.Shift))
            {
                e.SuppressKeyPress = true;
                Send();
            }
            // Escape closes the parent task pane if possible
            if (e.KeyCode == Keys.Escape)
            {
                e.SuppressKeyPress = true;
                _txtInput.Clear();
            }
        }

        private void Send()
        {
            if (_busy) return;
            string text = _txtInput.Text.Trim();
            if (text.Length == 0) return;

            _busy = true;
            _btnSend.Enabled = false;
            _lblStatus.Text = "Contacting Mistral...";

            AddBubble("You", text, true);
            _history.Add(new ChatMessage("user", text));
            if (_history.Count > MaxHistoryMessages)
            {
                _history.RemoveRange(0, _history.Count - MaxHistoryMessages);
            }
            _txtInput.Clear();

            // Read config + document context on the UI thread (Office OM is
            // main-thread only), then hand plain strings to the worker.
            Settings st = Settings.Load();
            string systemPrompt = st.SystemPrompt;
            string context = null;
            if (_chkContext.Checked)
            {
                _lblStatus.Text = "Reading " + DocNoun() + "...";
                context = GetDocumentContext();
            }

            string host = _host;
            string docNoun = DocNoun();
            ChatMessage[] snapshot = _history.ToArray();
            string baseUrl = st.BaseUrl, apiKey = st.ApiKey, model = st.Model;
            int timeout = st.TimeoutSeconds;

            // Capture the sync context for safe marshaling back to the UI thread.
            // This is safer than Control.BeginInvoke because the handle may not
            // exist yet during startup, or may be disposed during shutdown.
            SynchronizationContext ctx = _syncContext ?? SynchronizationContext.Current;

            ThreadPool.QueueUserWorkItem(delegate
            {
                string err;
                List<ChatMessage> msgs = new List<ChatMessage>();
                msgs.Add(new ChatMessage("system", BuildSystemPrompt(systemPrompt, context, host, docNoun)));
                foreach (ChatMessage m in snapshot) msgs.Add(m);
                string response = MistralClient.ChatCompletion(baseUrl, apiKey, model, msgs, timeout, out err);

                if (ctx != null)
                {
                    ctx.Post(delegate
                    {
                        OnResponseReceived(response, err);
                    }, null);
                }
                else
                {
                    // Fallback: try BeginInvoke if no SynchronizationContext
                    SafeBeginInvoke(delegate
                    {
                        OnResponseReceived(response, err);
                    });
                }
            });
        }

        private void OnResponseReceived(string response, string err)
        {
            if (response != null)
            {
                AddBubble("Mistral", response, false);
                _history.Add(new ChatMessage("assistant", response));
                _lastResponse = response;
                _lblStatus.Text = "Ready.";
            }
            else
            {
                AddBubble("Error", err ?? "Unknown error.", true, true);
                _lblStatus.Text = "Request failed.";
            }
            _busy = false;
            _btnSend.Enabled = true;
        }

        private string BuildSystemPrompt(string basePrompt, string context, string host, string docNoun)
        {
            if (string.IsNullOrEmpty(context)) return basePrompt;
            return basePrompt +
                "\r\n\r\nThe user is working in Microsoft " + host + ". Relevant content from the current " +
                docNoun + " follows between <context> tags. Use it when relevant.\r\n<context>\r\n" +
                context + "\r\n</context>";
        }

        private void ClearChat()
        {
            _history.Clear();
            _lastResponse = null;
            _chatPanel.Controls.Clear();
            _lblStatus.Text = "Ready.";
        }

        private void InsertLastResponse()
        {
            if (_lastResponse == null)
            {
                _lblStatus.Text = "No response to insert yet.";
                return;
            }
            try
            {
                InsertIntoDocument(_lastResponse);
                _lblStatus.Text = "Inserted into " + DocNoun() + ".";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not insert into the " + DocNoun() + ": " + ex.Message +
                    "\r\n\r\nTip: make sure a " + DocNoun() + " is open and a location is selected.",
                    "Mistral AI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void ShowSettings()
        {
            using (SettingsForm form = new SettingsForm())
            {
                form.ShowDialog(ParentForm != null ? ParentForm : SettingsForm.GetActiveOwner());
            }
        }

        #endregion

        #region Document context (Office OM via dynamic — main thread only)

        private string GetDocumentContext()
        {
            try
            {
                if (_appObj == null) return null;
                dynamic app = _appObj;
                if (_host == "Word") return GetWordContext(app);
                if (_host == "Excel") return GetExcelContext(app);
                if (_host == "PowerPoint") return GetPowerPointContext(app);
            }
            catch
            {
                // No document open / protected view / etc.
            }
            return null;
        }

        private string GetWordContext(dynamic app)
        {
            try
            {
                string text = Convert.ToString(app.ActiveDocument.Content.Text);
                if (text == null || text.Length == 0) return null;
                text = text.Replace('\r', '\n').Replace('\a', ' ');
                return Cap(text);
            }
            catch { return null; }
        }

        private string GetExcelContext(dynamic app)
        {
            try
            {
                dynamic range = app.ActiveSheet.UsedRange;
                int rows = Convert.ToInt32(range.Rows.Count);
                int cols = Convert.ToInt32(range.Columns.Count);
                if (rows <= 0 || cols <= 0) return null;

                StringBuilder sb = new StringBuilder();
                if (rows == 1 && cols == 1)
                {
                    object v = range.Value2;
                    if (v != null) sb.Append(Convert.ToString(v));
                }
                else
                {
                    object[,] vals = (object[,])range.Value2;
                    int maxRows = Math.Min(rows, 120);
                    int maxCols = Math.Min(cols, 40);
                    for (int r = 1; r <= maxRows; r++)
                    {
                        for (int c = 1; c <= maxCols; c++)
                        {
                            object v = vals[r, c];
                            if (v != null) sb.Append(Convert.ToString(v));
                            if (c < maxCols) sb.Append('\t');
                        }
                        sb.Append('\n');
                        if (sb.Length > MaxContextChars) break;
                    }
                    if (rows > maxRows) sb.Append("... (" + (rows - maxRows) + " more rows)");
                }
                return sb.Length == 0 ? null : Cap(sb.ToString());
            }
            catch { return null; }
        }

        private string GetPowerPointContext(dynamic app)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                dynamic slides = app.ActivePresentation.Slides;
                int count = Convert.ToInt32(slides.Count);
                int maxSlides = Math.Min(count, 80);
                for (int i = 1; i <= maxSlides; i++)
                {
                    dynamic slide = slides(i);
                    sb.Append("[Slide " + i + "]\n");
                    foreach (dynamic shape in slide.Shapes)
                    {
                        try
                        {
                            if (Convert.ToInt32(shape.TextFrame.HasText) == -1) // msoTrue
                            {
                                sb.Append(Convert.ToString(shape.TextFrame.TextRange.Text).Trim());
                                sb.Append('\n');
                            }
                        }
                        catch { /* shape without text frame */ }
                        if (sb.Length > MaxContextChars) break;
                    }
                    if (sb.Length > MaxContextChars) break;
                }
                if (count > maxSlides) sb.Append("... (" + (count - maxSlides) + " more slides)");
                return sb.Length == 0 ? null : Cap(sb.ToString());
            }
            catch { return null; }
        }

        private string Cap(string s)
        {
            s = s.Trim();
            if (s.Length <= MaxContextChars) return s;
            return s.Substring(0, MaxContextChars) + "\n... (truncated)";
        }

        private void InsertIntoDocument(string text)
        {
            dynamic app = _appObj;
            if (_host == "Word")
            {
                dynamic sel = app.Selection;
                sel.TypeText(text);
            }
            else if (_host == "Excel")
            {
                dynamic cell = app.ActiveCell;
                cell.Value2 = text;
            }
            else if (_host == "PowerPoint")
            {
                // 1 = msoTextOrientationHorizontal
                dynamic shape = app.ActiveWindow.View.Slide.Shapes.AddTextbox(1, 40, 60, 620, 380);
                shape.TextFrame.TextRange.Text = text;
            }
            else
            {
                throw new InvalidOperationException("Insert is not supported in this host.");
            }
        }

        #endregion

        #region Bubbles

        private void AddBubble(string role, string text, bool isUser)
        {
            AddBubble(role, text, isUser, false);
        }

        private void AddBubble(string role, string text, bool isUser, bool isError)
        {
            FlowLayoutPanel bubble = new FlowLayoutPanel();
            bubble.FlowDirection = FlowDirection.TopDown;
            bubble.WrapContents = false;
            bubble.AutoSize = true;
            bubble.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            bubble.Margin = new Padding(4, 4, 4, 6);
            bubble.Padding = new Padding(8);
            bubble.BackColor = isError ? Color.FromArgb(255, 235, 235)
                : (isUser ? Color.FromArgb(234, 234, 234) : Color.FromArgb(232, 240, 254));
            int width = Math.Max(_chatPanel.ClientSize.Width - 30, 160);
            bubble.MaximumSize = new Size(width, 0);

            Label roleLbl = new Label();
            roleLbl.Text = role;
            roleLbl.Font = new Font("Segoe UI", 8.5F, FontStyle.Bold);
            roleLbl.ForeColor = isError ? Color.Firebrick : (isUser ? Color.FromArgb(90, 90, 90) : Color.FromArgb(21, 66, 152));
            roleLbl.AutoSize = true;
            roleLbl.Margin = new Padding(0, 0, 0, 2);
            bubble.Controls.Add(roleLbl);

            Label txt = new Label();
            txt.Text = text;
            txt.Font = new Font("Segoe UI", 9F);
            txt.ForeColor = isError ? Color.Firebrick : SystemColors.ControlText;
            txt.AutoSize = true;
            txt.MaximumSize = new Size(width - 24, 0);
            bubble.Controls.Add(txt);

            _chatPanel.Controls.Add(bubble);
            _chatPanel.ScrollControlIntoView(bubble);
        }

        private void SafeBeginInvoke(MethodInvoker action)
        {
            try
            {
                if (IsHandleCreated && !IsDisposed)
                {
                    BeginInvoke(action);
                }
            }
            catch { /* control gone — app shutting down */ }
        }

        #endregion
    }
}
