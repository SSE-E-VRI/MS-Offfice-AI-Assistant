using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using MistralOfficeAddin.Core;

namespace MistralOfficeAddin.UI
{
    /// <summary>
    /// Hosts ChatSidebar as a child of the Office document window, docked on the right
    /// (Navigation-pane style). Used because Word 2010 CreateCTP cannot instantiate
    /// the .NET ActiveX control.
    /// </summary>
    public class OfficeDockedPane : Form
    {
        public const int DefaultPaneWidth = 380;
        private const int MinPaneWidth = 280;
        private const int MaxPaneWidth = 640;
        private const int SplitterWidth = 6;

        private readonly ChatSidebar _sidebar;
        private readonly Panel _splitter;
        private readonly ElementHost _elementHost;
        private IntPtr _parentHwnd = IntPtr.Zero;
        private IntPtr _documentViewHwnd = IntPtr.Zero;
        private Timer _layoutTimer;
        private int _paneWidth = DefaultPaneWidth;
        private bool _draggingSplitter;
        private int _dragStartX;
        private int _dragStartWidth;
        private bool _attached;
        private NativeWnd.EnumWindowsProc _enumChildren;

        public ChatSidebar Sidebar
        {
            get { return _sidebar; }
        }

        public OfficeDockedPane()
        {
            EnsureWpfApplication();

            this.Text = "AI Assistant";

            this.FormBorderStyle = FormBorderStyle.None;
            this.ControlBox = false;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.Manual;
            this.MinimumSize = new Size(MinPaneWidth, 200);
            this.Width = DefaultPaneWidth;
            this.BackColor = Color.FromArgb(0xE2, 0xE8, 0xF0);

            _splitter = new Panel();
            _splitter.Dock = DockStyle.Left;
            _splitter.Width = SplitterWidth;
            _splitter.Cursor = Cursors.VSplit;
            _splitter.BackColor = Color.FromArgb(0xCB, 0xD5, 0xE1);
            _splitter.MouseDown += Splitter_MouseDown;
            _splitter.MouseMove += Splitter_MouseMove;
            _splitter.MouseUp += Splitter_MouseUp;

            _elementHost = new ElementHost();
            _elementHost.Dock = DockStyle.Fill;
            _elementHost.TabStop = true;
            _sidebar = new ChatSidebar();
            _sidebar.PromptInputFocusRequested += Sidebar_PromptInputFocusRequested;
            _sidebar.PromptInputFocusLost += Sidebar_PromptInputFocusLost;
            _elementHost.Child = _sidebar;

            this.Controls.Add(_elementHost);
            this.Controls.Add(_splitter);
        }

        private int _lastDocLeft = -1;
        private int _lastDocTop = -1;
        private int _lastDocWidth = -1;
        private int _lastDocHeight = -1;
        private int _lastPaneX = -1;
        private int _lastPaneY = -1;
        private int _lastPaneW = -1;
        private int _lastPaneH = -1;

        public void InitializeHost(object appObj, string hostType)
        {
            if (_sidebar != null)
                _sidebar.InitializeHost(appObj, hostType);
        }

        public void ExecutePrompt(string prompt, string title)
        {
            if (_sidebar != null)
                _sidebar.ExecuteExternalPrompt(prompt, title);
        }

        public void StartNewChat()
        {
            if (_sidebar != null)
                _sidebar.StartNewChat(true);
        }

        public void FocusPromptInput()
        {
            try
            {
                ActivatePromptKeyboardRouting();
                if (_sidebar != null)
                    _sidebar.FocusPromptInput();
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("OfficeDockedPane.FocusPromptInput failed: {0}", ex.Message));
            }
        }

        private void Sidebar_PromptInputFocusRequested(object sender, EventArgs e)
        {
            if (this.IsHandleCreated)
                this.BeginInvoke(new MethodInvoker(ActivatePromptKeyboardRouting));
        }

        private void Sidebar_PromptInputFocusLost(object sender, EventArgs e)
        {
            // Keyboard focus left the sidebar entirely (e.g. user clicked the sheet).
            // Hand the keys back to Excel.
            _routeKeysToPrompt = false;
        }

        private NativeWnd.RECT _originalDocRect = default(NativeWnd.RECT);
        private bool _hasOriginalDocRect = false;

        public bool AttachToHostWindow(IntPtr parentHwnd)
        {
            if (parentHwnd == IntPtr.Zero || !NativeWnd.IsWindow(parentHwnd))
            {
                Logger.Warn("OfficeDockedPane: host HWND is invalid.");
                return false;
            }

            _parentHwnd = parentHwnd;

            if (!this.IsHandleCreated)
                this.CreateControl();
            if (!this.IsHandleCreated)
                this.Show();

            NativeWnd.SetParent(this.Handle, parentHwnd);

            long style = NativeWnd.GetWindowLongPtr(this.Handle, NativeWnd.GWL_STYLE).ToInt64();
            style |= NativeWnd.WS_CHILD | NativeWnd.WS_VISIBLE | NativeWnd.WS_CLIPSIBLINGS | NativeWnd.WS_CLIPCHILDREN;
            style &= ~(NativeWnd.WS_POPUP | NativeWnd.WS_CAPTION | NativeWnd.WS_SYSMENU | NativeWnd.WS_THICKFRAME | NativeWnd.WS_MINIMIZEBOX | NativeWnd.WS_MAXIMIZEBOX);
            NativeWnd.SetWindowLongPtr(this.Handle, NativeWnd.GWL_STYLE, new IntPtr(style));

            long ex = NativeWnd.GetWindowLongPtr(this.Handle, NativeWnd.GWL_EXSTYLE).ToInt64();
            ex &= ~NativeWnd.WS_EX_APPWINDOW;
            ex &= ~NativeWnd.WS_EX_DLGMODALFRAME;
            NativeWnd.SetWindowLongPtr(this.Handle, NativeWnd.GWL_EXSTYLE, new IntPtr(ex));

            _documentViewHwnd = FindLargestDirectChild(parentHwnd, this.Handle);
            if (_documentViewHwnd != IntPtr.Zero && NativeWnd.IsWindow(_documentViewHwnd))
            {
                NativeWnd.RECT r;
                if (NativeWnd.GetWindowRect(_documentViewHwnd, out r))
                {
                    NativeWnd.POINT tl = new NativeWnd.POINT { X = r.Left, Y = r.Top };
                    NativeWnd.POINT br = new NativeWnd.POINT { X = r.Right, Y = r.Bottom };
                    NativeWnd.ScreenToClient(parentHwnd, ref tl);
                    NativeWnd.ScreenToClient(parentHwnd, ref br);
                    _originalDocRect = new NativeWnd.RECT { Left = tl.X, Top = tl.Y, Right = br.X, Bottom = br.Y };
                    _hasOriginalDocRect = true;
                }
            }

            _attached = true;

            if (_layoutTimer == null)
            {
                _layoutTimer = new Timer();
                _layoutTimer.Interval = 250;
                _layoutTimer.Tick += delegate { ApplyLayout(); };
            }
            _layoutTimer.Start();

            ApplyLayout();
            this.Show();
            NativeWnd.SetWindowPos(this.Handle, NativeWnd.HWND_TOP, 0, 0, 0, 0,
                NativeWnd.SWP_NOMOVE | NativeWnd.SWP_NOSIZE | NativeWnd.SWP_NOACTIVATE | NativeWnd.SWP_FRAMECHANGED | NativeWnd.SWP_SHOWWINDOW);

            FocusPromptInput();
            InstallPromptKeyboardHook();
            Logger.Info("OfficeDockedPane attached to host window (right dock) with standard WPF keyboard routing.");
            return true;
        }

        public void DetachAndRestore()
        {
            UninstallPromptKeyboardHook();
            _routeKeysToPrompt = false;
            _promptInputHwnd = IntPtr.Zero;
            _focusRetryCount = 0;
            try
            {
                RestoreDocumentView();
            }
            catch { }

            if (_layoutTimer != null)
            {
                try { _layoutTimer.Stop(); } catch { }
            }

            _attached = false;
            _parentHwnd = IntPtr.Zero;
            _documentViewHwnd = IntPtr.Zero;
            _hasOriginalDocRect = false;
            _lastDocLeft = -1;
            _lastDocTop = -1;
            _lastDocWidth = -1;
            _lastDocHeight = -1;
            _lastPaneX = -1;
            _lastPaneY = -1;
            _lastPaneW = -1;
            _lastPaneH = -1;
        }

        private void ApplyLayout()
        {
            if (!_attached || _parentHwnd == IntPtr.Zero || !NativeWnd.IsWindow(_parentHwnd))
                return;
            if (!this.IsHandleCreated)
                return;

            NativeWnd.RECT client;
            if (!NativeWnd.GetClientRect(_parentHwnd, out client))
                return;

            int clientW = client.Right - client.Left;
            int clientH = client.Bottom - client.Top;
            if (clientW < MinPaneWidth + 80 || clientH < 80)
                return;

            int paneW = _paneWidth;
            if (paneW > clientW - 80) paneW = clientW - 80;
            if (paneW < MinPaneWidth) paneW = MinPaneWidth;
            _paneWidth = paneW;

            int paneX = clientW - paneW;
            int paneY = 0;
            int paneH = clientH;

            if (_documentViewHwnd != IntPtr.Zero && NativeWnd.IsWindow(_documentViewHwnd)
                && NativeWnd.GetParent(_documentViewHwnd) == _parentHwnd)
            {
                NativeWnd.RECT docRect;
                NativeWnd.GetWindowRect(_documentViewHwnd, out docRect);
                NativeWnd.POINT topLeft = new NativeWnd.POINT { X = docRect.Left, Y = docRect.Top };
                NativeWnd.ScreenToClient(_parentHwnd, ref topLeft);
                int docLeft = topLeft.X;
                int docTop = topLeft.Y;
                int docBottom = docTop + (docRect.Bottom - docRect.Top);
                if (docLeft < 0) docLeft = 0;
                if (docTop < 0) docTop = 0;
                int docWidth = paneX - docLeft;
                int docHeight = clientH - docTop;
                if (docBottom > 0 && docBottom < clientH)
                    docHeight = docBottom - docTop;
                if (docWidth < 60) docWidth = 60;
                if (docHeight < 60) docHeight = 60;

                if (docLeft != _lastDocLeft || docTop != _lastDocTop || docWidth != _lastDocWidth || docHeight != _lastDocHeight)
                {
                    _lastDocLeft = docLeft;
                    _lastDocTop = docTop;
                    _lastDocWidth = docWidth;
                    _lastDocHeight = docHeight;
                    NativeWnd.MoveWindow(_documentViewHwnd, docLeft, docTop, docWidth, docHeight, true);
                }

                // Position docked assistant pane directly below the Ribbon / Formula Bar
                paneY = docTop;
                paneH = docHeight;
            }

            if (paneX != _lastPaneX || paneY != _lastPaneY || paneW != _lastPaneW || paneH != _lastPaneH)
            {
                _lastPaneX = paneX;
                _lastPaneY = paneY;
                _lastPaneW = paneW;
                _lastPaneH = paneH;
                NativeWnd.MoveWindow(this.Handle, paneX, paneY, paneW, paneH, true);
            }
        }

        private void RestoreDocumentView()
        {
            if (_parentHwnd == IntPtr.Zero || _documentViewHwnd == IntPtr.Zero)
                return;
            if (!NativeWnd.IsWindow(_parentHwnd) || !NativeWnd.IsWindow(_documentViewHwnd))
                return;

            if (_hasOriginalDocRect)
            {
                int origW = _originalDocRect.Right - _originalDocRect.Left;
                int origH = _originalDocRect.Bottom - _originalDocRect.Top;
                if (origW > 0 && origH > 0)
                {
                    NativeWnd.MoveWindow(_documentViewHwnd, _originalDocRect.Left, _originalDocRect.Top, origW, origH, true);
                    return;
                }
            }

            NativeWnd.RECT client;
            if (!NativeWnd.GetClientRect(_parentHwnd, out client))
                return;

            NativeWnd.RECT docRect;
            NativeWnd.GetWindowRect(_documentViewHwnd, out docRect);
            NativeWnd.POINT topLeft = new NativeWnd.POINT { X = docRect.Left, Y = docRect.Top };
            NativeWnd.ScreenToClient(_parentHwnd, ref topLeft);
            int docLeft = topLeft.X < 0 ? 0 : topLeft.X;
            int docTop = topLeft.Y < 0 ? 0 : topLeft.Y;
            NativeWnd.MoveWindow(
                _documentViewHwnd,
                docLeft,
                docTop,
                (client.Right - client.Left) - docLeft,
                (client.Bottom - client.Top) - docTop,
                true);
        }

        private IntPtr FindLargestDirectChild(IntPtr parent, IntPtr exclude)
        {
            IntPtr best = IntPtr.Zero;
            int bestArea = 0;
            _enumChildren = delegate(IntPtr hWnd, IntPtr lParam)
            {
                if (hWnd == exclude) return true;
                if (NativeWnd.GetParent(hWnd) != parent) return true;

                NativeWnd.RECT r;
                if (!NativeWnd.GetWindowRect(hWnd, out r)) return true;

                var name = new StringBuilder(64);
                NativeWnd.GetClassName(hWnd, name, 63);
                string cls = name.ToString();
                if (string.Equals(cls, "_WwG", StringComparison.Ordinal) ||
                    string.Equals(cls, "_WwB", StringComparison.Ordinal))
                {
                    best = hWnd;
                    bestArea = int.MaxValue;
                    return false;
                }

                int area = (r.Right - r.Left) * (r.Bottom - r.Top);
                if (area > bestArea)
                {
                    bestArea = area;
                    best = hWnd;
                }
                return true;
            };
            NativeWnd.EnumChildWindows(parent, _enumChildren, IntPtr.Zero);
            return best;
        }

        private void Splitter_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            _draggingSplitter = true;
            _dragStartX = Control.MousePosition.X;
            _dragStartWidth = _paneWidth;
            _splitter.Capture = true;
        }

        private void Splitter_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_draggingSplitter) return;
            int delta = _dragStartX - Control.MousePosition.X;
            int next = _dragStartWidth + delta;
            if (next < MinPaneWidth) next = MinPaneWidth;
            if (next > MaxPaneWidth) next = MaxPaneWidth;
            _paneWidth = next;
            ApplyLayout();
        }

        private void Splitter_MouseUp(object sender, MouseEventArgs e)
        {
            _draggingSplitter = false;
            _splitter.Capture = false;
        }

        // Excel 2010's native loop can retain focus on the worksheet even after a WPF child
        // TextBox is clicked (GetFocus() stays on EXCEL7).  When the sidebar owns keyboard
        // focus this hook consumes each queued key message, dispatches it to the WPF HwndSource,
        // and rewrites it to WM_NULL so Excel's pump never feeds the active cell.  When the
        // sidebar does not own focus, the message is left alone and Excel handles it normally.
        private bool _routeKeysToPrompt;
        private IntPtr _promptInputHwnd = IntPtr.Zero;
        private IntPtr _keyboardHook = IntPtr.Zero;
        private NativeWnd.HookProc _keyboardHookProc;
        private bool _dispatchingHook;
        private int _focusRetryCount;

        private void ActivatePromptKeyboardRouting()
        {
            if (_sidebar == null) return;

            IntPtr promptWindow = _sidebar.GetPromptInputWindowHandle();
            if (promptWindow == IntPtr.Zero || !NativeWnd.IsWindow(promptWindow))
            {
                // Attach runs before the WPF visual is connected; retry until the HwndSource exists.
                if (_focusRetryCount < 10 && this.IsHandleCreated)
                {
                    _focusRetryCount++;
                    this.BeginInvoke(new MethodInvoker(ActivatePromptKeyboardRouting));
                }
                return;
            }

            _focusRetryCount = 0;
            _promptInputHwnd = promptWindow;
            _routeKeysToPrompt = true;
            NativeWnd.SetFocus(promptWindow);
        }

        private bool IsInPane(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero || !this.IsHandleCreated) return false;
            return hwnd == this.Handle ||
                   (_elementHost != null && _elementHost.IsHandleCreated && hwnd == _elementHost.Handle) ||
                   NativeWnd.IsChild(this.Handle, hwnd);
        }

        private void InstallPromptKeyboardHook()
        {
            if (_keyboardHook != IntPtr.Zero) return;
            try
            {
                _keyboardHookProc = new NativeWnd.HookProc(PromptKeyboardHookCallback);
                _keyboardHook = NativeWnd.SetWindowsHookEx(
                    NativeWnd.WH_GETMESSAGE,
                    _keyboardHookProc,
                    IntPtr.Zero,
                    NativeWnd.GetCurrentThreadId());
                if (_keyboardHook == IntPtr.Zero)
                    Logger.Warn("OfficeDockedPane: could not install the Excel prompt keyboard bridge.");
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("OfficeDockedPane: prompt keyboard bridge setup failed: {0}", ex.Message));
            }
        }

        private void UninstallPromptKeyboardHook()
        {
            if (_keyboardHook == IntPtr.Zero) return;
            try { NativeWnd.UnhookWindowsHookEx(_keyboardHook); }
            catch { }
            _keyboardHook = IntPtr.Zero;
            _keyboardHookProc = null;
        }

        private IntPtr PromptKeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam.ToInt64() == NativeWnd.PM_REMOVE && !_dispatchingHook)
            {
                try
                {
                    NativeWnd.MSG message = (NativeWnd.MSG)Marshal.PtrToStructure(lParam, typeof(NativeWnd.MSG));
                    if (message.message == NativeWnd.WM_LBUTTONDOWN ||
                        message.message == NativeWnd.WM_RBUTTONDOWN ||
                        message.message == NativeWnd.WM_MBUTTONDOWN ||
                        message.message == NativeWnd.WM_NCLBUTTONDOWN)
                    {
                        if (IsInPane(message.hwnd))
                        {
                            // Any in-pane click re-arms key routing so pane controls (combo, list, etc.) work.
                            // If Excel kept native focus (GetFocus still on EXCEL7), force it back to the
                            // WPF HwndSource so WPF processes the input.
                            _routeKeysToPrompt = true;
                            IntPtr currentFocus = NativeWnd.GetFocus();
                            if (_promptInputHwnd != IntPtr.Zero && NativeWnd.IsWindow(_promptInputHwnd) &&
                                (currentFocus == IntPtr.Zero || !IsInPane(currentFocus)))
                            {
                                NativeWnd.SetFocus(_promptInputHwnd);
                            }
                        }
                        else
                        {
                            // Click outside the assistant returns keyboard ownership to Excel
                            _routeKeysToPrompt = false;
                        }
                    }
                    else if (message.message >= NativeWnd.WM_KEYFIRST &&
                             message.message <= NativeWnd.WM_KEYLAST &&
                             _routeKeysToPrompt &&
                             _sidebar != null && _sidebar.IsPromptKeyboardFocused &&
                             _promptInputHwnd != IntPtr.Zero && NativeWnd.IsWindow(_promptInputHwnd))
                    {
                        // The sidebar owns keyboard focus: route the key to WPF (the HwndSource
                        // delivers it to whichever pane element is focused) and consume the message
                        // so Excel's pump never starts in-cell editing.
                        message.hwnd = _promptInputHwnd;
                        NativeWnd.TranslateMessage(ref message);
                        _dispatchingHook = true;
                        try
                        {
                            NativeWnd.DispatchMessage(ref message);
                        }
                        finally
                        {
                            _dispatchingHook = false;
                        }
                        message.message = NativeWnd.WM_NULL;
                        Marshal.StructureToPtr(message, lParam, false);
                    }
                }
                catch (Exception hookEx)
                {
                    Logger.Warn(string.Format("OfficeDockedPane: keyboard hook callback error: {0}", hookEx.Message));
                }
            }

            return NativeWnd.CallNextHookEx(_keyboardHook, nCode, wParam, lParam);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            UninstallPromptKeyboardHook();
            DetachAndRestore();
            base.OnFormClosing(e);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                UninstallPromptKeyboardHook();
                DetachAndRestore();
            }
            base.Dispose(disposing);
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
                Logger.Warn(string.Format("OfficeDockedPane WPF Application init: {0}", ex.Message));
            }
        }

        public static IntPtr ResolveHostHwnd(object appObj)
        {
            if (appObj == null) return IntPtr.Zero;
            try
            {
                dynamic app = appObj;
                try
                {
                    dynamic window = app.ActiveWindow;
                    if (window != null)
                    {
                        long hwnd = Convert.ToInt64(window.Hwnd);
                        if (hwnd != 0) return new IntPtr(hwnd);
                    }
                }
                catch { }

                try
                {
                    long hwnd = Convert.ToInt64(app.Hwnd);
                    if (hwnd != 0) return new IntPtr(hwnd);
                }
                catch { }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("ResolveHostHwnd failed: {0}", ex.Message));
            }
            return IntPtr.Zero;
        }
    }

    internal static class NativeWnd
    {
        public const int GWL_STYLE = -16;
        public const int GWL_EXSTYLE = -20;
        public const int WM_SIZE = 0x0005;
        public const int WM_WINDOWPOSCHANGED = 0x0047;
        public const int WS_CHILD = 0x40000000;
        public const int WS_VISIBLE = 0x10000000;
        public const int WS_CLIPSIBLINGS = 0x04000000;
        public const int WS_CLIPCHILDREN = 0x02000000;
        public const int WS_CAPTION = 0x00C00000;
        public const int WS_SYSMENU = 0x00080000;
        public const int WS_THICKFRAME = 0x00040000;
        public const int WS_MINIMIZEBOX = 0x00020000;
        public const int WS_MAXIMIZEBOX = 0x00010000;
        public const int WS_POPUP = unchecked((int)0x80000000);
        public const int WS_EX_APPWINDOW = 0x00040000;
        public const int WS_EX_DLGMODALFRAME = 0x00000001;
        public const int SWP_NOMOVE = 0x0002;
        public const int SWP_NOSIZE = 0x0001;
        public const int SWP_NOZORDER = 0x0004;
        public const int SWP_NOACTIVATE = 0x0010;
        public const int SWP_FRAMECHANGED = 0x0020;
        public const int SWP_SHOWWINDOW = 0x0040;
        public static readonly IntPtr HWND_TOP = new IntPtr(0);

        public const int WH_GETMESSAGE = 3;
        public const int PM_REMOVE = 1;
        public const int WM_NULL = 0x0000;
        public const int WM_KEYFIRST = 0x0100;
        public const int WM_KEYLAST = 0x0109;
        public const int WM_LBUTTONDOWN = 0x0201;
        public const int WM_RBUTTONDOWN = 0x0204;
        public const int WM_MBUTTONDOWN = 0x0207;
        public const int WM_NCLBUTTONDOWN = 0x00A1;

        public delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        public static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        public static extern bool IsChild(IntPtr hWndParent, IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        public static IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex)
        {
            if (IntPtr.Size == 8)
                return GetWindowLongPtr64(hWnd, nIndex);
            else
                return GetWindowLong32(hWnd, nIndex);
        }

        public static IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong)
        {
            if (IntPtr.Size == 8)
                return SetWindowLongPtr64(hWnd, nIndex, dwNewLong);
            else
                return SetWindowLong32(hWnd, nIndex, dwNewLong);
        }

        [DllImport("user32.dll", EntryPoint = "GetWindowLong", SetLastError = true)]
        private static extern IntPtr GetWindowLong32(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr", SetLastError = true)]
        private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "SetWindowLong", SetLastError = true)]
        private static extern IntPtr SetWindowLong32(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr", SetLastError = true)]
        private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        [DllImport("user32.dll")]
        public static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("user32.dll")]
        public static extern IntPtr GetParent(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool ScreenToClient(IntPtr hWnd, ref POINT lpPoint);

        [DllImport("user32.dll")]
        public static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll")]
        public static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        public static extern IntPtr SetFocus(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr GetFocus();

        [DllImport("user32.dll")]
        public static extern bool TranslateMessage([In] ref MSG lpMsg);

        [DllImport("user32.dll")]
        public static extern IntPtr DispatchMessage([In] ref MSG lpmsg);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MSG
        {
            public IntPtr hwnd;
            public uint message;
            public IntPtr wParam;
            public IntPtr lParam;
            public uint time;
            public POINT pt;
        }

    }
}
