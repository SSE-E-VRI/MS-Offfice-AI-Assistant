using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using Microsoft.Win32;
using MSOfficeAIAssistant.Core;
using MSOfficeAIAssistant.UI;

namespace MSOfficeAIAssistant.Addin
{
    #region COM Interfaces for Task Panes & Object Safety

    /// <summary>
    /// Microsoft.Office.Core.ICustomTaskPaneConsumer
    /// </summary>
    [ComImport]
    [Guid("000C033E-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICustomTaskPaneConsumer
    {
        [DispId(1)]
        void CTPFactoryAvailable([MarshalAs(UnmanagedType.IDispatch)] object CTPFactoryInst);
    }

    /// <summary>
    /// Microsoft.Office.Core.ICTPFactory
    /// </summary>
    [ComImport]
    [Guid("000C033D-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ICTPFactory
    {
        [DispId(1)]
        object CreateCTP(
            [MarshalAs(UnmanagedType.BStr)] string CTPAxID,
            [MarshalAs(UnmanagedType.BStr)] string CTPTitle,
            [MarshalAs(UnmanagedType.Struct)] object CTPParentWindow);
    }

    /// <summary>
    /// Microsoft.Office.Core._CustomTaskPane
    /// </summary>
    [ComImport]
    [Guid("000C033B-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface _CustomTaskPane
    {
        [DispId(1)]
        string Title { get; }
        [DispId(2)]
        object Application { get; }
        [DispId(3)]
        object Window { get; }
        [DispId(4)]
        bool Visible { get; set; }
        [DispId(5)]
        object ContentControl { get; }
        [DispId(6)]
        int Height { get; set; }
        [DispId(7)]
        int Width { get; set; }
        [DispId(8)]
        int DockPosition { get; set; }
        [DispId(9)]
        int DockPositionRestrict { get; set; }
        [DispId(10)]
        void Delete();
    }

    /// <summary>
    /// IObjectSafety interface required by Office for ActiveX hosting in Custom Task Panes.
    /// </summary>
    [ComImport]
    [Guid("CB5BDC81-93C1-11CF-8F20-00805F2CD064")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IObjectSafety
    {
        [PreserveSig]
        int GetInterfaceSafetyOptions(
            ref Guid riid,
            out int pdwSupportedOptions,
            out int pdwEnabledOptions);

        [PreserveSig]
        int SetInterfaceSafetyOptions(
            ref Guid riid,
            int dwOptionSetMask,
            int dwEnabledOptions);
    }

    #endregion

    /// <summary>
    /// ActiveX wrapper control hosting the WPF ChatSidebar inside an ElementHost.
    /// IMPORTANT: The WPF sidebar is created LAZILY on first use (not in constructor)
    /// to prevent COM thread deadlock during Office startup.
    /// </summary>
    [ComVisible(true)]
    [Guid("9B3C7624-5A1D-4C5E-8C9B-12D3E4F5A6B7")]
    [ProgId("MSOfficeAIAssistant.TaskPaneControl")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class TaskPaneControl : UserControl, IObjectSafety
    {
        public const string ControlClsid = "{9B3C7624-5A1D-4C5E-8C9B-12D3E4F5A6B7}";
        public const string ControlProgId = "MSOfficeAIAssistant.TaskPaneControl";

        private const int INTERFACESAFE_FOR_UNTRUSTED_CALLER = 0x00000001;
        private const int INTERFACESAFE_FOR_UNTRUSTED_DATA = 0x00000002;
        private const int S_OK = 0;

        // Sidebar is created lazily to avoid blocking the COM startup thread
        private ElementHost _elementHost;
        private ChatSidebar _wpfSidebar;
        private bool _sidebarInitialized;
        private readonly object _initLock = new object();

        // Stored for deferred initialization
        private object _pendingAppObj;
        private string _pendingHostType;

        public TaskPaneControl()
        {
            // Minimal constructor - no WPF creation here.
            // WPF init happens lazily in EnsureSidebarCreated().
            this.SuspendLayout();
            this.Name = "TaskPaneControl";
            this.Size = new System.Drawing.Size(380, 700);
            this.BackColor = System.Drawing.Color.FromArgb(0xFA, 0xFA, 0xFA);
            this.TabStop = true;
            this.ResumeLayout(false);

            // Schedule sidebar creation on the UI message pump once the handle is created
            this.HandleCreated += OnHandleCreated;
        }

        [ComRegisterFunction]
        public static void RegisterControl(Type t)
        {
            if (t == null || t.GUID != new Guid(ControlClsid)) return;
            WriteActiveXRegistration(Registry.CurrentUser);
            try { WriteActiveXRegistration(Registry.LocalMachine); }
            catch { }
        }

        [ComUnregisterFunction]
        public static void UnregisterControl(Type t)
        {
            if (t == null || t.GUID != new Guid(ControlClsid)) return;
            RemoveActiveXRegistration(Registry.CurrentUser);
            try { RemoveActiveXRegistration(Registry.LocalMachine); }
            catch { }
        }

        internal static void WriteActiveXRegistration(RegistryKey hive)
        {
            string keyPath = @"Software\Classes\CLSID\" + ControlClsid;
            using (RegistryKey key = hive.CreateSubKey(keyPath))
            {
                if (key == null) return;
                key.CreateSubKey("Control").Close();
                using (RegistryKey ms = key.CreateSubKey("MiscStatus"))
                {
                    if (ms != null) ms.SetValue("", "131473");
                }
                using (RegistryKey cats = key.CreateSubKey("Implemented Categories"))
                {
                    if (cats == null) return;
                    cats.CreateSubKey("{7DD95801-9882-11CF-9FA9-00AA006C42C4}").Close();
                    cats.CreateSubKey("{7DD95802-9882-11CF-9FA9-00AA006C42C4}").Close();
                    cats.CreateSubKey("{40FC6ED4-2438-11CF-A3DB-080036F12502}").Close();
                }
            }
        }

        private static void RemoveActiveXRegistration(RegistryKey hive)
        {
            string keyPath = @"Software\Classes\CLSID\" + ControlClsid;
            using (RegistryKey key = hive.OpenSubKey(keyPath, true))
            {
                if (key == null) return;
                try { key.DeleteSubKey("Control", false); } catch { }
                try { key.DeleteSubKeyTree("Implemented Categories"); } catch { }
            }
        }

        private void OnHandleCreated(object sender, EventArgs e)
        {
            // Post deferred WPF creation so it runs after COM returns
            this.BeginInvoke(new MethodInvoker(EnsureSidebarCreated));
        }

        private static void EnsureWpfApplication()
        {
            if (System.Windows.Application.Current != null) return;
            try
            {
                var app = new System.Windows.Application();
                app.ShutdownMode = System.Windows.ShutdownMode.OnExplicitShutdown;
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("Could not create WPF Application: {0}", ex.Message));
            }
        }

        private void EnsureSidebarCreated()
        {
            if (_sidebarInitialized) return;
            lock (_initLock)
            {
                if (_sidebarInitialized) return;

                try
                {
                    EnsureWpfApplication();

                    _elementHost = new ElementHost();
                    _elementHost.Dock = DockStyle.Fill;
                    _elementHost.Name = "elementHost";
                    _elementHost.TabStop = true;

                    _wpfSidebar = new ChatSidebar();
                    _elementHost.Child = _wpfSidebar;

                    this.SuspendLayout();
                    this.Controls.Add(_elementHost);
                    this.ResumeLayout(true);

                    _sidebarInitialized = true;
                    Logger.Info("TaskPaneControl: WPF ChatSidebar created successfully (lazy init).");

                    // Apply any deferred host initialization
                    if (_pendingAppObj != null)
                    {
                        ApplyHostInit(_pendingAppObj, _pendingHostType);
                        _pendingAppObj = null;
                        _pendingHostType = null;
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error("TaskPaneControl: Failed to create WPF sidebar", ex);
                }
            }
        }

        protected override void WndProc(ref Message m)
        {
            // Let Windows Forms and ElementHost forward focus and keyboard messages to the
            // hosted WPF controls. Forcing focus to ElementHost here prevents Excel from
            // giving the WPF TextBox a real editable keyboard focus.
            base.WndProc(ref m);
        }

        public void InitializeHost(object appObj, string hostType)
        {
            if (!_sidebarInitialized || _wpfSidebar == null)
            {
                // Store for deferred application when sidebar is ready
                _pendingAppObj = appObj;
                _pendingHostType = hostType;
                return;
            }

            ApplyHostInit(appObj, hostType);
        }

        private void ApplyHostInit(object appObj, string hostType)
        {
            try
            {
                if (_wpfSidebar != null)
                {
                    _wpfSidebar.InitializeHost(appObj, hostType);
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("TaskPaneControl.ApplyHostInit failed: {0}", ex.Message));
            }
        }

        public void ExecutePrompt(string prompt, string title)
        {
            EnsureSidebarCreated();
            try
            {
                if (_wpfSidebar != null)
                {
                    _wpfSidebar.ExecuteExternalPrompt(prompt, title);
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("TaskPaneControl.ExecutePrompt failed: {0}", ex.Message));
            }
        }

        public void StartNewChat()
        {
            EnsureSidebarCreated();
            try
            {
                if (_wpfSidebar != null)
                {
                    _wpfSidebar.StartNewChat(true);
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("TaskPaneControl.StartNewChat failed: {0}", ex.Message));
            }
        }

        public void ReloadConfiguredProvider()
        {
            try
            {
                if (_wpfSidebar != null)
                {
                    _wpfSidebar.ReloadConfiguredProvider();
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("TaskPaneControl.ReloadConfiguredProvider failed: {0}", ex.Message));
            }
        }

        #region IObjectSafety Implementation

        public int GetInterfaceSafetyOptions(ref Guid riid, out int pdwSupportedOptions, out int pdwEnabledOptions)
        {
            pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER | INTERFACESAFE_FOR_UNTRUSTED_DATA;
            pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER | INTERFACESAFE_FOR_UNTRUSTED_DATA;
            return S_OK;
        }

        public int SetInterfaceSafetyOptions(ref Guid riid, int dwOptionSetMask, int dwEnabledOptions)
        {
            return S_OK;
        }

        #endregion

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                try
                {
                    _wpfSidebar = null;
                    if (_elementHost != null)
                    {
                        _elementHost.Child = null;
                        _elementHost.Dispose();
                        _elementHost = null;
                    }
                }
                catch { }
            }
            base.Dispose(disposing);
        }
    }

    /// <summary>
    /// Manages the lifecycle of Office Custom Task Panes
    /// </summary>
    public class CustomTaskPaneManager
    {
        private object _ctpFactoryObj;
        private dynamic _customTaskPane;
        private TaskPaneControl _taskPaneControl;
        private ChatFloatingWindow _floatingWindow;
        private OfficeDockedPane _dockedPane;
        private object _appObj;
        private string _hostType = "Office";

        public bool IsVisible
        {
            get
            {
                try
                {
                    if (_customTaskPane != null)
                        return (bool)_customTaskPane.Visible;
                    if (_dockedPane != null && !_dockedPane.IsDisposed)
                        return _dockedPane.Visible;
                    if (_floatingWindow != null && _floatingWindow.IsLoaded)
                        return _floatingWindow.IsVisible;
                    return false;
                }
                catch
                {
                    return false;
                }
            }
        }

        public void SetFactory(object factory)
        {
            _ctpFactoryObj = factory;
            Logger.Info(string.Format("CTP Factory set: {0}", factory != null ? factory.GetType().FullName : "null"));
        }

        public void SetHost(object appObj, string hostType)
        {
            _appObj = appObj;
            _hostType = hostType ?? "Office";
            if (_taskPaneControl != null)
            {
                _taskPaneControl.InitializeHost(_appObj, _hostType);
            }
            if (_dockedPane != null && !_dockedPane.IsDisposed)
            {
                _dockedPane.InitializeHost(_appObj, _hostType);
            }
            if (_floatingWindow != null)
            {
                _floatingWindow.InitializeHost(_appObj, _hostType);
            }
        }

        public void TogglePane()
        {
            if (_customTaskPane != null)
            {
                try
                {
                    bool current = (bool)_customTaskPane.Visible;
                    _customTaskPane.Visible = !current;
                    if (!current)
                    {
                        DockPaneToRight(_customTaskPane);
                    }
                    return;
                }
                catch (Exception ex)
                {
                    Logger.Warn(string.Format("Failed to toggle task pane visibility: {0}", ex.Message));
                    _customTaskPane = null;
                    _taskPaneControl = null;
                }
            }
            else if (_dockedPane != null && !_dockedPane.IsDisposed)
            {
                if (_dockedPane.Visible)
                    HideDockedPane();
                else
                    ShowDockedPane();
                return;
            }
            else if (_floatingWindow != null && _floatingWindow.IsLoaded)
            {
                if (_floatingWindow.IsVisible)
                {
                    _floatingWindow.Hide();
                }
                else
                {
                    _floatingWindow.Show();
                    _floatingWindow.Activate();
                    _floatingWindow.FocusPromptInput();
                }
                return;
            }

            CreateAndShowPane();
        }

        public void ShowPane()
        {
            if (_customTaskPane != null)
            {
                try
                {
                    _customTaskPane.Visible = true;
                    DockPaneToRight(_customTaskPane);
                    return;
                }
                catch
                {
                    _customTaskPane = null;
                    _taskPaneControl = null;
                }
            }
            else if (_dockedPane != null && !_dockedPane.IsDisposed)
            {
                ShowDockedPane();
                return;
            }
            else if (_floatingWindow != null && _floatingWindow.IsLoaded)
            {
                _floatingWindow.Show();
                _floatingWindow.Activate();
                _floatingWindow.FocusPromptInput();
                return;
            }

            CreateAndShowPane();
        }

        public void ExecutePrompt(string prompt, string title)
        {
            ShowPane();
            if (_taskPaneControl != null)
            {
                _taskPaneControl.ExecutePrompt(prompt, title);
            }
            else if (_dockedPane != null && !_dockedPane.IsDisposed)
            {
                _dockedPane.ExecutePrompt(prompt, title);
            }
            else if (_floatingWindow != null)
            {
                _floatingWindow.ExecutePrompt(prompt, title);
            }
        }

        public void StartNewChat()
        {
            ShowPane();
            if (_taskPaneControl != null)
            {
                _taskPaneControl.StartNewChat();
            }
            else if (_dockedPane != null && !_dockedPane.IsDisposed)
            {
                _dockedPane.StartNewChat();
            }
            else if (_floatingWindow != null)
            {
                _floatingWindow.StartNewChat();
            }
        }

        public void ReloadConfiguredProvider()
        {
            try
            {
                if (_taskPaneControl != null)
                {
                    _taskPaneControl.ReloadConfiguredProvider();
                }
                if (_dockedPane != null && !_dockedPane.IsDisposed && _dockedPane.Sidebar != null)
                {
                    _dockedPane.Sidebar.ReloadConfiguredProvider();
                }
                if (_floatingWindow != null && _floatingWindow.Sidebar != null)
                {
                    _floatingWindow.Sidebar.ReloadConfiguredProvider();
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("CustomTaskPaneManager.ReloadConfiguredProvider failed: {0}", ex.Message));
            }
        }

        /// <summary>
        /// Opens the embedded User Manual in a modeless window owned by the Office host window
        /// so the manual stays in front. Works without any API key.
        /// </summary>
        public void OpenUserManual()
        {
            try
            {
                IntPtr hostWindow = OfficeDockedPane.ResolveHostHwnd(_appObj);
                HelpWindow.ShowHelp(hostWindow);
            }
            catch (Exception ex)
            {
                Logger.Error("CustomTaskPaneManager.OpenUserManual failed", ex);
            }
        }

        private void CreateAndShowPane()
        {
            // Excel 2010 hosts the managed ActiveX task-pane control but does not reliably
            // paint the ElementHost child (it appears as a blank white pane). Use the native
            // child-window docked host first so Excel retains a right-side assistant pane.
            // A top-level WPF window remains only as the final fallback.
            if (RequiresFloatingExcel2010Host())
            {
                Logger.Info("Using manually docked AI chat host for Excel 2010 compatibility.");
                if (ShowDockedPane())
                    return;

                Logger.Warn("Excel 2010 docked host was unavailable; using floating AI chat fallback.");
                ShowFloatingWindowFallback();
                return;
            }

            // Prefer Office's native Custom Task Pane whenever its factory is available.
            // Excel supplies this factory and its native pane correctly routes keyboard focus
            // into ElementHost. The manually parented docked pane is a fallback only.
            TryHealActiveXRegistration();

            if (_ctpFactoryObj != null)
            {
                try
                {
                    Logger.Info("Creating Custom Task Pane with ProgId MSOfficeAIAssistant.TaskPaneControl...");
                    dynamic factory = _ctpFactoryObj;
                    object parentWindow = ResolveParentWindow();

                    _customTaskPane = TryCreateCtp(factory, TaskPaneControl.ControlProgId, parentWindow);
                    if (_customTaskPane == null)
                    {
                        _customTaskPane = TryCreateCtp(factory, TaskPaneControl.ControlClsid, parentWindow);
                    }
                    if (_customTaskPane == null && parentWindow != Type.Missing)
                    {
                        _customTaskPane = TryCreateCtp(factory, TaskPaneControl.ControlProgId, Type.Missing);
                    }

                    if (_customTaskPane != null)
                    {
                        DockPaneToRight(_customTaskPane);

                        try
                        {
                            object content = _customTaskPane.ContentControl;
                            _taskPaneControl = content as TaskPaneControl;
                            if (_taskPaneControl != null)
                            {
                                _taskPaneControl.InitializeHost(_appObj, _hostType);
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.Warn(string.Format("Failed to obtain TaskPaneControl from ContentControl: {0}", ex.Message));
                        }

                        _customTaskPane.Visible = true;
                        try { _customTaskPane.Width = 380; } catch { }
                        Logger.Info("Custom Task Pane created and docked to the right.");
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error("Failed to create Custom Task Pane, falling back to floating window", ex);
                }
            }
            else
            {
                Logger.Warn("CTP Factory is null; trying docked-pane fallback.");
            }

            // Use the manually docked pane only when Office cannot create a native CTP.
            if (ShowDockedPane())
                return;

            // Final fallback to a standalone WPF window.
            ShowFloatingWindowFallback();
        }

        private bool RequiresFloatingExcel2010Host()
        {
            if (!string.Equals(_hostType, "Excel", StringComparison.OrdinalIgnoreCase)) return false;
            try
            {
                dynamic app = _appObj;
                string version = app != null ? Convert.ToString(app.Version) : string.Empty;
                return version != null && version.StartsWith("14.", StringComparison.Ordinal);
            }
            catch
            {
                // The compatibility fallback is limited to an explicitly identified Excel host;
                // if the version cannot be read, keep the normal native-pane path.
                return false;
            }
        }

        private bool ShowDockedPane()
        {
            try
            {
                IntPtr hwnd = OfficeDockedPane.ResolveHostHwnd(_appObj);
                if (hwnd == IntPtr.Zero)
                {
                    Logger.Warn("ShowDockedPane: could not resolve host window HWND.");
                    return false;
                }

                if (_dockedPane == null || _dockedPane.IsDisposed)
                {
                    _dockedPane = new OfficeDockedPane();
                    _dockedPane.InitializeHost(_appObj, _hostType);
                    _dockedPane.FormClosed += delegate
                    {
                        _dockedPane = null;
                    };
                    if (!_dockedPane.AttachToHostWindow(hwnd))
                    {
                        try { _dockedPane.Dispose(); } catch { }
                        _dockedPane = null;
                        return false;
                    }
                }
                else
                {
                    _dockedPane.Show();
                    _dockedPane.AttachToHostWindow(hwnd);
                }

                _dockedPane.FocusPromptInput();
                Logger.Info("Docked assistant pane shown on the right of the document window.");
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error("ShowDockedPane failed", ex);
                return false;
            }
        }

        private void HideDockedPane()
        {
            if (_dockedPane == null || _dockedPane.IsDisposed)
                return;
            try
            {
                _dockedPane.DetachAndRestore();
                _dockedPane.Hide();
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("HideDockedPane failed: {0}", ex.Message));
            }
        }

        private static object TryCreateCtp(dynamic factory, string axId, object parentWindow)
        {
            try
            {
                object pane = factory.CreateCTP(axId, "AI Assistant", parentWindow);
                if (pane != null)
                {
                    Logger.Info(string.Format("CreateCTP succeeded with '{0}'.", axId));
                    return pane;
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("CreateCTP('{0}') failed: {1}", axId, ex.Message));
            }
            return null;
        }

        private object ResolveParentWindow()
        {
            try
            {
                if (_appObj == null) return Type.Missing;
                dynamic app = _appObj;
                object window = app.ActiveWindow;
                if (window != null) return window;
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("Could not resolve ActiveWindow for task pane: {0}", ex.Message));
            }
            return Type.Missing;
        }

        private static void DockPaneToRight(dynamic pane)
        {
            if (pane == null) return;
            const int msoCTPDockPositionRight = 2;
            try { pane.DockPosition = msoCTPDockPositionRight; }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("Failed to set DockPosition=Right: {0}", ex.Message));
            }
            try { pane.Width = 380; } catch { }
        }

        private static void TryHealActiveXRegistration()
        {
            try
            {
                TaskPaneControl.WriteActiveXRegistration(Registry.CurrentUser);
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("HKCU ActiveX registration heal failed: {0}", ex.Message));
            }
            try
            {
                TaskPaneControl.WriteActiveXRegistration(Registry.LocalMachine);
                Logger.Info("Task pane ActiveX control registered under HKLM (required by Office 2010 CreateCTP).");
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("HKLM ActiveX registration skipped (admin usually required): {0}", ex.Message));
            }
        }

        private void ShowFloatingWindowFallback()
        {
            try
            {
                if (_floatingWindow == null || !_floatingWindow.IsLoaded)
                {
                    _floatingWindow = new ChatFloatingWindow();
                    _floatingWindow.InitializeHost(_appObj, _hostType);
                    _floatingWindow.Closed += (s, e) => { _floatingWindow = null; };
                }
                _floatingWindow.Show();
                _floatingWindow.Activate();
                _floatingWindow.FocusPromptInput();
            }
            catch (Exception ex)
            {
                Logger.Error("Failed to show fallback floating window", ex);
                MessageBox.Show(
                    string.Format("Could not open AI Chat window: {0}", ex.Message),
                    "Mistral AI",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public void Cleanup()
        {
            try
            {
                if (_customTaskPane != null)
                {
                    try { _customTaskPane.Delete(); } catch { }
                    _customTaskPane = null;
                }
                if (_dockedPane != null)
                {
                    try
                    {
                        _dockedPane.DetachAndRestore();
                        _dockedPane.Close();
                    }
                    catch { }
                    _dockedPane = null;
                }
                if (_floatingWindow != null)
                {
                    try { _floatingWindow.Close(); } catch { }
                    _floatingWindow = null;
                }
            }
            catch { }
            _taskPaneControl = null;
            _ctpFactoryObj = null;
            _appObj = null;
        }
    }
}
