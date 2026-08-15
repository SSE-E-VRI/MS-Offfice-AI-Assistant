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

            this.Text = "Mistral AI Assistant";
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

            var host = new ElementHost();
            host.Dock = DockStyle.Fill;
            _sidebar = new ChatSidebar();
            host.Child = _sidebar;

            this.Controls.Add(host);
            this.Controls.Add(_splitter);
        }

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

            Logger.Info("OfficeDockedPane attached to host window (right dock).");
            return true;
        }

        public void DetachAndRestore()
        {
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
                NativeWnd.MoveWindow(_documentViewHwnd, docLeft, docTop, docWidth, docHeight, true);
            }

            NativeWnd.MoveWindow(this.Handle, paneX, 0, paneW, clientH, true);
        }

        private void RestoreDocumentView()
        {
            if (_parentHwnd == IntPtr.Zero || _documentViewHwnd == IntPtr.Zero)
                return;
            if (!NativeWnd.IsWindow(_parentHwnd) || !NativeWnd.IsWindow(_documentViewHwnd))
                return;

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

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            DetachAndRestore();
            base.OnFormClosing(e);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                DetachAndRestore();
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
                        int hwnd = Convert.ToInt32(window.Hwnd);
                        if (hwnd != 0) return new IntPtr(hwnd);
                    }
                }
                catch { }

                try
                {
                    int hwnd = Convert.ToInt32(app.Hwnd);
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

        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr", SetLastError = true)]
        public static extern IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr", SetLastError = true)]
        public static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

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
    }
}
