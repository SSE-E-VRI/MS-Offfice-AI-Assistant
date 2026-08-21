using System;
using System.Runtime.InteropServices;
using MSOfficeAIAssistant.Core;

namespace MSOfficeAIAssistant.Tests
{
    public static class ComResilienceTests
    {
        public static void RunAll()
        {
            TestOleMessageFilterComInterfaceIdentity();
            TestOleMessageFilterHandlerReturns();
            TestOleMessageFilterRetryLogic();
            TestOleMessageFilterRegistrationLifecycle();
            TestSafeOfficeProbeSuccess();
            TestSafeOfficeProbeComExceptionFallback();
            TestSafeOfficeProbeGeneralExceptionFallback();
            TestSafeOfficeProbeTryExecute();
        }

        private static void TestOleMessageFilterComInterfaceIdentity()
        {
            // Verify canonical IID_IMessageFilter: 00000016-0000-0000-C000-000000000046
            Guid expectedIid = new Guid("00000016-0000-0000-C000-000000000046");
            Guid actualIid = typeof(IOleMessageFilter).GUID;
            Assert(actualIid == expectedIid,
                string.Format("IOleMessageFilter GUID mismatch! Expected canonical IID_IMessageFilter {0}, got {1}",
                    expectedIid, actualIid));

            // Verify CCW creation and COM QueryInterface for IID_IMessageFilter
            var filter = new OleMessageFilter();
            IntPtr pCom = Marshal.GetComInterfaceForObject(filter, typeof(IOleMessageFilter));
            Assert(pCom != IntPtr.Zero, "CCW must successfully create and expose COM interface pointer for IOleMessageFilter");
            Marshal.Release(pCom);
        }

        private static void TestOleMessageFilterHandlerReturns()
        {
            var filter = new OleMessageFilter();
            int handleCall = filter.HandleInComingCall(0, IntPtr.Zero, 0, IntPtr.Zero);
            Assert(handleCall == 0, "HandleInComingCall should return SERVERCALL_ISHANDLED (0), got " + handleCall);

            int pending = filter.MessagePending(IntPtr.Zero, 0, 0);
            Assert(pending == 2, "MessagePending should return PENDINGMSG_WAITDEFPROCESS (2), got " + pending);
        }

        private static void TestOleMessageFilterRetryLogic()
        {
            var filter = new OleMessageFilter();

            // SERVERCALL_RETRYLATER (2) under 10s -> retry delay 150ms
            int retryUnder10s = filter.RetryRejectedCall(IntPtr.Zero, dwTickCount: 1500, dwRejectType: 2);
            Assert(retryUnder10s == 150, "Expected retry delay 150ms for SERVERCALL_RETRYLATER under 10s, got " + retryUnder10s);

            // SERVERCALL_RETRYLATER (2) at or over 10s -> cancel (-1)
            int retryOver10s = filter.RetryRejectedCall(IntPtr.Zero, dwTickCount: 10500, dwRejectType: 2);
            Assert(retryOver10s == -1, "Expected -1 (cancel) for SERVERCALL_RETRYLATER over 10s, got " + retryOver10s);

            // Other reject types (e.g. SERVERCALL_REJECTED = 1) -> cancel (-1)
            int rejected = filter.RetryRejectedCall(IntPtr.Zero, dwTickCount: 500, dwRejectType: 1);
            Assert(rejected == -1, "Expected -1 for SERVERCALL_REJECTED, got " + rejected);
        }

        private static void TestOleMessageFilterRegistrationLifecycle()
        {
            Exception threadEx = null;
            var staThread = new System.Threading.Thread(new System.Threading.ThreadStart(() =>
            {
                try
                {
                    Assert(!OleMessageFilter.IsRegistered, "Should not be registered initially");

                    OleMessageFilter.Register();
                    Assert(OleMessageFilter.IsRegistered, "Should be registered after Register() on STA thread");

                    // Idempotent register
                    OleMessageFilter.Register();
                    Assert(OleMessageFilter.IsRegistered, "Should remain registered");

                    OleMessageFilter.Revoke();
                    Assert(!OleMessageFilter.IsRegistered, "Should not be registered after Revoke()");

                    // Idempotent revoke
                    OleMessageFilter.Revoke();
                    Assert(!OleMessageFilter.IsRegistered, "Should remain unregistered");
                }
                catch (Exception ex)
                {
                    threadEx = ex;
                }
            }));

            staThread.SetApartmentState(System.Threading.ApartmentState.STA);
            staThread.Start();
            staThread.Join();

            if (threadEx != null)
            {
                throw threadEx;
            }
        }

        private static void TestSafeOfficeProbeSuccess()
        {
            int value = SafeOfficeProbe.Probe(() => 42, fallback: 0, probeName: "SuccessProbe");
            Assert(value == 42, "SafeOfficeProbe should return function result on success, got " + value);
        }

        private static void TestSafeOfficeProbeComExceptionFallback()
        {
            // Simulate COM 0x800AC472 (VBA_E_IGNORE) or similar
            string result = SafeOfficeProbe.Probe<string>(
                () => { throw new COMException("Excel is in cell edit mode", unchecked((int)0x800AC472)); },
                fallback: "FallbackValue",
                probeName: "ExcelCellEditProbe");

            Assert(result == "FallbackValue", "SafeOfficeProbe should return fallback on COMException, got " + result);
        }

        private static void TestSafeOfficeProbeGeneralExceptionFallback()
        {
            bool result = SafeOfficeProbe.Probe<bool>(
                () => { throw new InvalidOperationException("API mismatch"); },
                fallback: false,
                probeName: "MismatchProbe");

            Assert(!result, "SafeOfficeProbe should return fallback false on Exception");
        }

        private static void TestSafeOfficeProbeTryExecute()
        {
            bool executed = false;
            bool success = SafeOfficeProbe.TryExecute(() => { executed = true; }, "ActionProbe");
            Assert(success && executed, "TryExecute should succeed when action executes");

            bool failed = SafeOfficeProbe.TryExecute(() => { throw new Exception("Action error"); }, "FailingActionProbe");
            Assert(!failed, "TryExecute should return false on exception without throwing");
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition)
            {
                throw new Exception(message);
            }
        }
    }
}
