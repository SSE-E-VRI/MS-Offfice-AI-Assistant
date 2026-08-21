using System;
using System.Runtime.InteropServices;

namespace MSOfficeAIAssistant.Core
{
    [ComImport]
    [Guid("00000016-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IOleMessageFilter
    {
        [PreserveSig]
        int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);

        [PreserveSig]
        int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);

        [PreserveSig]
        int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
    }

    /// <summary>
    /// Implements IOleMessageFilter to handle Office COM busy rejections (e.g. Excel in-cell edit 0x800AC472,
    /// VBA_E_IGNORE, RPC_E_SERVERCALL_RETRYLATER, or modal dialogs).
    /// Registers on the STA thread during add-in connection and restores the previous filter on disconnect.
    /// </summary>
    public class OleMessageFilter : IOleMessageFilter
    {
        private const int SERVERCALL_ISHANDLED = 0;
        private const int SERVERCALL_REJECTED = 1;
        private const int SERVERCALL_RETRYLATER = 2;
        private const int PENDINGMSG_WAITDEFPROCESS = 2;

        [DllImport("Ole32.dll", ExactSpelling = true, PreserveSig = true)]
        private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);

        [ThreadStatic]
        private static IOleMessageFilter _previousFilter;

        [ThreadStatic]
        private static bool _isRegisteredOnThread;

        public static bool IsRegistered
        {
            get { return _isRegisteredOnThread; }
        }

        public static void Register()
        {
            if (_isRegisteredOnThread) return;

            try
            {
                IOleMessageFilter newFilter = new OleMessageFilter();
                IOleMessageFilter previous;
                int hr = CoRegisterMessageFilter(newFilter, out previous);
                if (hr >= 0)
                {
                    _previousFilter = previous;
                    _isRegisteredOnThread = true;
                    Logger.Info("OleMessageFilter registered successfully on STA thread.");
                }
                else
                {
                    Logger.Warn(string.Format("CoRegisterMessageFilter failed with HRESULT: 0x{0:X8}", hr));
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("OleMessageFilter.Register failed: {0}", ex.Message));
            }
        }

        public static void Revoke()
        {
            if (!_isRegisteredOnThread) return;

            try
            {
                IOleMessageFilter dummy;
                CoRegisterMessageFilter(_previousFilter, out dummy);
                _previousFilter = null;
                _isRegisteredOnThread = false;
                Logger.Info("OleMessageFilter revoked and previous filter restored.");
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("OleMessageFilter.Revoke failed: {0}", ex.Message));
            }
        }

        public int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
        {
            return SERVERCALL_ISHANDLED;
        }

        public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
        {
            // dwRejectType: 2 == SERVERCALL_RETRYLATER (Excel formula edit, modal dialog, busy computation)
            if (dwRejectType == SERVERCALL_RETRYLATER)
            {
                // Retry for up to 10 seconds (10,000 ms)
                if (dwTickCount < 10000)
                {
                    // Return retry interval in ms (150ms delay between COM retries)
                    return 150;
                }
                Logger.Warn("OleMessageFilter: COM call rejected by host exceeded retry timeout (10s).");
            }
            // -1: Cancel the rejected call immediately, raising RPC_E_SERVERCALL_RETRYLATER to caller
            return -1;
        }

        public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
        {
            return PENDINGMSG_WAITDEFPROCESS;
        }
    }
}
