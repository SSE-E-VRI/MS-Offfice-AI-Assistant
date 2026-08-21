using System;

namespace MSOfficeAIAssistant.Hosts
{
    /// <summary>
    /// Structured result returned by host controllers for mutation and inspection operations.
    /// Provides explicit success/failure indication, error codes, diagnostic messages, and optional return values
    /// so the Phase C Verification Engine can distinguish between clean execution, harmless no-ops, and silent COM rejections.
    /// </summary>
    public class HostOperationResult
    {
        public bool Success { get; private set; }
        public string ErrorMessage { get; private set; }
        public int ErrorCode { get; private set; }
        public object Value { get; private set; }
        public string TargetLocation { get; private set; }

        public static HostOperationResult Ok(object value = null, string targetLocation = null)
        {
            return new HostOperationResult
            {
                Success = true,
                Value = value,
                TargetLocation = targetLocation,
                ErrorMessage = null,
                ErrorCode = 0
            };
        }

        public static HostOperationResult Failed(string message, int errorCode = 0, string targetLocation = null)
        {
            return new HostOperationResult
            {
                Success = false,
                ErrorMessage = message ?? "Operation failed",
                ErrorCode = errorCode,
                TargetLocation = targetLocation,
                Value = null
            };
        }

        public static HostOperationResult FromException(Exception ex, string operationName = null, string targetLocation = null)
        {
            if (ex == null) return Failed("Unknown error", 0, targetLocation);
            int hr = 0;
            try
            {
                hr = System.Runtime.InteropServices.Marshal.GetHRForException(ex);
            }
            catch { }

            string opPrefix = !string.IsNullOrEmpty(operationName) ? (operationName + ": ") : string.Empty;
            return Failed(string.Format("{0}{1}", opPrefix, ex.Message), hr, targetLocation);
        }
    }
}
