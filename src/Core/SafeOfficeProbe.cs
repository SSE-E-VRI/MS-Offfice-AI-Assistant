using System;
using System.Runtime.InteropServices;

namespace MSOfficeAIAssistant.Core
{
    /// <summary>
    /// Typed helper for non-mutating Office COM property reads, optional feature discovery,
    /// and version-compatibility probes (Office 2010 vs 2016 vs 365 differences).
    /// Provides structured diagnostics while isolating non-mutating capability checks.
    /// </summary>
    public static class SafeOfficeProbe
    {
        public static T Probe<T>(Func<T> probeFunc, T fallback = default(T), string probeName = null)
        {
            if (probeFunc == null) return fallback;
            try
            {
                return probeFunc();
            }
            catch (COMException comEx)
            {
                if (probeName != null)
                {
                    Logger.Debug(string.Format("SafeOfficeProbe [{0}] COMException (0x{1:X8}): {2}", probeName, comEx.ErrorCode, comEx.Message));
                }
                return fallback;
            }
            catch (Exception ex)
            {
                if (probeName != null)
                {
                    Logger.Debug(string.Format("SafeOfficeProbe [{0}] Exception: {1}", probeName, ex.Message));
                }
                return fallback;
            }
        }

        public static bool TryExecute(Action action, string actionName = null)
        {
            if (action == null) return false;
            try
            {
                action();
                return true;
            }
            catch (COMException comEx)
            {
                if (actionName != null)
                {
                    Logger.Warn(string.Format("SafeOfficeProbe [{0}] COMException (0x{1:X8}): {2}", actionName, comEx.ErrorCode, comEx.Message));
                }
                return false;
            }
            catch (Exception ex)
            {
                if (actionName != null)
                {
                    Logger.Warn(string.Format("SafeOfficeProbe [{0}] Exception: {1}", actionName, ex.Message));
                }
                return false;
            }
        }
    }
}
