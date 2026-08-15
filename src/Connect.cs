using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace MistralOfficeAddin
{
    /// <summary>
    /// The COM add-in entry point. One class serves Word, Excel and PowerPoint;
    /// the host is detected in OnConnection. Registered under
    /// HKCU\Software\Microsoft\Office\{Word,Excel,PowerPoint}\Addins\MistralAI.Connect
    /// with LoadBehavior=3 (load at startup).
    /// </summary>
    [Guid("2F8D4B61-7C3E-4A59-9B2D-6E1F0A3C5E78")]
    [ProgId("MistralAI.Connect")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility, ICustomTaskPaneConsumer
    {
        private object _appObj;                 // host Application RCW (accessed via dynamic)
        private string _host = "Office";        // Word | Excel | PowerPoint
        private object _ctpFactory;             // ICTPFactory as object — use dynamic for calls
        private dynamic _pane;                  // _CustomTaskPane via dynamic
        private TaskPaneControl _control;
        private SynchronizationContext _syncContext;

        private static string LogPath
        {
            get { return Path.Combine(Path.GetTempPath(), "MistralAddinLog.txt"); }
        }

        private static void Log(string msg)
        {
            try
            {
                string line = DateTime.Now.ToString("HH:mm:ss.fff") + " [" + Thread.CurrentThread.ManagedThreadId + "] " + msg;
                File.AppendAllText(LogPath, line + Environment.NewLine);
            }
            catch { }
        }

        #region IDTExtensibility2

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            Log("OnConnection ENTERED");
            try
            {
                Log("Application type=" + (Application != null ? Application.GetType().FullName : "null"));
                _appObj = Application;

                // Capture UI SynchronizationContext for safe cross-thread marshaling
                _syncContext = SynchronizationContext.Current;

                dynamic app = Application;
                string name = Convert.ToString(app.Name);
                Log("App.Name=" + name);

                if (name.IndexOf("Word", StringComparison.OrdinalIgnoreCase) >= 0) _host = "Word";
                else if (name.IndexOf("Excel", StringComparison.OrdinalIgnoreCase) >= 0) _host = "Excel";
                else if (name.IndexOf("PowerPoint", StringComparison.OrdinalIgnoreCase) >= 0) _host = "PowerPoint";

                Log("Host=" + _host + " ConnectMode=" + ConnectMode);
                Log("OnConnection OK");
            }
            catch (Exception ex)
            {
                Log("OnConnection ERROR: " + ex.GetType().Name + ": " + ex.Message);
                if (ex.InnerException != null)
                    Log("  Inner: " + ex.InnerException.GetType().Name + ": " + ex.InnerException.Message);
                Log("  Stack: " + ex.StackTrace);
                throw;  // re-throw so Office sees the failure
            }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            Log("OnDisconnection");
            try
            {
                if (_pane != null)
                {
                    try { _pane.Delete(); } catch { }
                }
            }
            catch { }
            _control = null;
            _pane = null;
            _ctpFactory = null;
            _appObj = null;
            _syncContext = null;
        }

        public void OnAddInsUpdate(ref Array custom) { Log("OnAddInsUpdate"); }
        public void OnStartupComplete(ref Array custom) { Log("OnStartupComplete"); }
        public void OnBeginShutdown(ref Array custom) { Log("OnBeginShutdown"); }

        #endregion

        #region IRibbonExtensibility

        public string GetCustomUI(string RibbonID)
        {
            Log("GetCustomUI RibbonID=" + RibbonID);
            // Same ribbon works for Microsoft.Word.Word, Microsoft.Excel.Workbook
            // and Microsoft.PowerPoint.Presentation.
            return RibbonXml.Ribbon;
        }

        #endregion

        #region ICustomTaskPaneConsumer

        public void CTPFactoryAvailable(object CTPFactoryInst)
        {
            Log("CTPFactoryAvailable received factory: " + (CTPFactoryInst != null ? CTPFactoryInst.GetType().FullName : "null"));
            _ctpFactory = CTPFactoryInst;
        }

        #endregion

        #region Ribbon callbacks (invoked by name via IDispatch)

        public void OnChatButtonClick(object control)
        {
            try
            {
                ToggleChatPane();
            }
            catch (Exception ex)
            {
                Log("OnChatButtonClick ERROR: " + ex.GetType().Name + ": " + ex.Message);
                MessageBox.Show("Could not open the chat pane: " + ex.Message, "Mistral AI",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnSettingsButtonClick(object control)
        {
            try
            {
                using (SettingsForm form = new SettingsForm())
                {
                    form.ShowDialog(SettingsForm.GetActiveOwner());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not open settings: " + ex.Message, "Mistral AI",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        private void ToggleChatPane()
        {
            if (_pane != null)
            {
                try
                {
                    _pane.Visible = !(bool)_pane.Visible;
                }
                catch (Exception ex)
                {
                    Log("ToggleChatPane visibility toggle failed: " + ex.Message);
                    // Pane may have been disposed; recreate
                    _pane = null;
                    _control = null;
                }
                if (_pane != null) return;
            }

            if (_ctpFactory == null)
            {
                MessageBox.Show("This host did not provide a task pane factory. The chat window is unavailable in " +
                    _host + " on this machine.\n\nTry restarting the Office application.",
                    "Mistral AI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Log("Creating CTP with ProgID MistralAI.ChatPane");
            try
            {
                // Use dynamic to call CreateCTP — the factory is received as
                // object from CTPFactoryAvailable and the return type is a
                // COM _CustomTaskPane.  dynamic handles all IDispatch dispatch.
                dynamic factory = _ctpFactory;
                _pane = factory.CreateCTP("MistralAI.ChatPane", "Mistral AI", Type.Missing);
                Log("CreateCTP returned: " + (_pane != null ? _pane.GetType().FullName : "null"));
            }
            catch (Exception ex)
            {
                Log("CreateCTP FAILED: " + ex.GetType().Name + ": " + ex.Message);
                MessageBox.Show("Could not create the chat pane: " + ex.Message +
                    "\n\nMake sure register.cmd was run successfully.",
                    "Mistral AI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try { _pane.DockPosition = 2; } catch { }   // msoCTPDockPositionRight
            try { _pane.Width = 380; } catch { }

            // The ContentControl property returns the hosted ActiveX control
            try
            {
                object content = _pane.ContentControl;
                _control = content as TaskPaneControl;
                Log("ContentControl type=" + (content != null ? content.GetType().FullName : "null") +
                    " cast=" + (_control != null));
            }
            catch (Exception ex)
            {
                Log("ContentControl access failed: " + ex.Message);
            }

            if (_control != null) _control.Initialize(_appObj, _host);
            _pane.Visible = true;
        }
    }
}
