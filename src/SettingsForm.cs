using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace MistralOfficeAddin
{
    /// <summary>
    /// Settings dialog: API key (DPAPI-encrypted on save), model, base URL,
    /// timeout and system prompt. Includes a background "Test Connection"
    /// call to GET /v1/models.
    /// </summary>
    public class SettingsForm : Form
    {
        private static readonly string[] KnownModels = new string[]
        {
            "mistral-small-latest",
            "open-mistral-nemo",
            "mistral-large-latest"
        };

        private TextBox _txtApiKey;
        private ComboBox _cboModel;
        private TextBox _txtBaseUrl;
        private NumericUpDown _numTimeout;
        private TextBox _txtSystemPrompt;
        private CheckBox _chkContext;
        private Button _btnTest;
        private Button _btnSave;
        private Button _btnCancel;

        public SettingsForm()
        {
            Text = "Mistral AI Settings";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(470, 470);
            Font = new Font("Segoe UI", 9F);

            BuildUi();
            LoadSettings();
        }

        private void BuildUi()
        {
            Label l;

            l = NewLabel("Mistral API key:", 14, 16);
            Controls.Add(l);
            _txtApiKey = new TextBox();
            _txtApiKey.Location = new Point(14, 36);
            _txtApiKey.Size = new Size(440, 23);
            Controls.Add(_txtApiKey);

            l = NewLabel("Model (free tier: mistral-small-latest / open-mistral-nemo):", 14, 68);
            Controls.Add(l);
            _cboModel = new ComboBox();
            _cboModel.DropDownStyle = ComboBoxStyle.DropDown; // editable
            _cboModel.Location = new Point(14, 88);
            _cboModel.Size = new Size(440, 23);
            foreach (string m in KnownModels) _cboModel.Items.Add(m);
            Controls.Add(_cboModel);

            l = NewLabel("API base URL:", 14, 120);
            Controls.Add(l);
            _txtBaseUrl = new TextBox();
            _txtBaseUrl.Location = new Point(14, 140);
            _txtBaseUrl.Size = new Size(440, 23);
            Controls.Add(_txtBaseUrl);

            l = NewLabel("Request timeout (seconds):", 14, 172);
            Controls.Add(l);
            _numTimeout = new NumericUpDown();
            _numTimeout.Location = new Point(14, 192);
            _numTimeout.Size = new Size(80, 23);
            _numTimeout.Minimum = 5;
            _numTimeout.Maximum = 300;
            Controls.Add(_numTimeout);

            _chkContext = new CheckBox();
            _chkContext.Text = "Include document context by default";
            _chkContext.Location = new Point(110, 194);
            _chkContext.AutoSize = true;
            Controls.Add(_chkContext);

            l = NewLabel("System prompt:", 14, 226);
            Controls.Add(l);
            _txtSystemPrompt = new TextBox();
            _txtSystemPrompt.Multiline = true;
            _txtSystemPrompt.ScrollBars = ScrollBars.Vertical;
            _txtSystemPrompt.Location = new Point(14, 246);
            _txtSystemPrompt.Size = new Size(440, 90);
            Controls.Add(_txtSystemPrompt);

            _btnTest = new Button();
            _btnTest.Text = "Test Connection";
            _btnTest.Location = new Point(14, 354);
            _btnTest.Size = new Size(120, 30);
            _btnTest.Click += delegate { TestConnection(); };
            Controls.Add(_btnTest);

            _btnSave = new Button();
            _btnSave.Text = "Save";
            _btnSave.Location = new Point(268, 354);
            _btnSave.Size = new Size(88, 30);
            _btnSave.Click += delegate { SaveAndClose(); };
            Controls.Add(_btnSave);

            _btnCancel = new Button();
            _btnCancel.Text = "Cancel";
            _btnCancel.Location = new Point(364, 354);
            _btnCancel.Size = new Size(88, 30);
            _btnCancel.DialogResult = DialogResult.Cancel;
            Controls.Add(_btnCancel);
        }

        private Label NewLabel(string text, int x, int y)
        {
            Label l = new Label();
            l.Text = text;
            l.Location = new Point(x, y);
            l.AutoSize = true;
            return l;
        }

        private void LoadSettings()
        {
            Settings st = Settings.Load();
            _txtApiKey.Text = st.ApiKey;
            _cboModel.Text = st.Model;
            _txtBaseUrl.Text = st.BaseUrl;
            _numTimeout.Value = st.TimeoutSeconds;
            _txtSystemPrompt.Text = st.SystemPrompt;
            _chkContext.Checked = st.IncludeContextByDefault;
        }

        private Settings Collect()
        {
            Settings st = new Settings();
            st.ApiKey = _txtApiKey.Text.Trim();
            st.Model = _cboModel.Text.Trim();
            if (st.Model.Length == 0) st.Model = "mistral-small-latest";
            st.BaseUrl = _txtBaseUrl.Text.Trim();
            if (st.BaseUrl.Length == 0) st.BaseUrl = "https://api.mistral.ai/v1";
            st.TimeoutSeconds = (int)_numTimeout.Value;
            st.SystemPrompt = _txtSystemPrompt.Text;
            st.IncludeContextByDefault = _chkContext.Checked;
            return st;
        }

        private void SaveAndClose()
        {
            try
            {
                Collect().Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Could not save settings: " + ex.Message, "Mistral AI",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            DialogResult = DialogResult.OK;
            Close();
        }

        private void TestConnection()
        {
            // Test with what is currently typed, before saving.
            Settings st = Collect();
            _btnTest.Enabled = false;
            _btnTest.Text = "Testing...";
            string baseUrl = st.BaseUrl, apiKey = st.ApiKey;
            int timeout = st.TimeoutSeconds;

            ThreadPool.QueueUserWorkItem(delegate
            {
                string err;
                string result = MistralClient.TestConnection(baseUrl, apiKey, timeout, out err);
                try
                {
                    BeginInvoke((MethodInvoker)delegate
                    {
                        _btnTest.Enabled = true;
                        _btnTest.Text = "Test Connection";
                        if (result != null)
                        {
                            MessageBox.Show(this, result, "Mistral AI \u2014 Connection OK",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show(this, err ?? "Unknown error.", "Mistral AI \u2014 Connection failed",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    });
                }
                catch { /* dialog closed */ }
            });
        }

        /// <summary>
        /// Returns a wrapper around the current foreground window so modal
        /// dialogs parent correctly onto the Office window.
        /// </summary>
        public static IWin32Window GetActiveOwner()
        {
            IntPtr h = NativeMethods.GetActiveWindow();
            if (h != IntPtr.Zero) return new WindowHandleWrapper(h);
            return null;
        }

        private class WindowHandleWrapper : IWin32Window
        {
            private readonly IntPtr _handle;
            public WindowHandleWrapper(IntPtr handle) { _handle = handle; }
            public IntPtr Handle { get { return _handle; } }
        }
    }

    internal static class NativeMethods
    {
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern IntPtr GetActiveWindow();
    }
}
