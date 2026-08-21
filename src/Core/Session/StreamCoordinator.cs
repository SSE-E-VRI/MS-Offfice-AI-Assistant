using System;
using System.Text;
using System.Threading;

namespace MSOfficeAIAssistant.Core.Session
{
    /// <summary>
    /// Coordinates asynchronous token streaming, delta accumulation, UI update throttling,
    /// cancellation, and completion guards to prevent race conditions during UI updates.
    /// Thread-safe and independent of WPF Dispatcher.
    /// </summary>
    public class StreamCoordinator
    {
        private readonly object _syncRoot = new object();
        private readonly StringBuilder _accumulator = new StringBuilder();
        private int _deltaCount;
        private bool _isStreamFinished;
        private bool _isSending;
        private CancellationTokenSource _cts;

        public bool IsSending
        {
            get
            {
                lock (_syncRoot)
                {
                    return _isSending;
                }
            }
            set
            {
                lock (_syncRoot)
                {
                    _isSending = value;
                }
            }
        }

        public bool IsStreamFinished
        {
            get
            {
                lock (_syncRoot)
                {
                    return _isStreamFinished;
                }
            }
        }

        public CancellationTokenSource BeginStream()
        {
            lock (_syncRoot)
            {
                if (_cts != null)
                {
                    try { _cts.Cancel(); } catch { }
                    try { _cts.Dispose(); } catch { }
                }

                _accumulator.Length = 0;
                _deltaCount = 0;
                _isStreamFinished = false;
                _isSending = true;
                _cts = new CancellationTokenSource();
                return _cts;
            }
        }

        public void AccumulateDelta(string delta, Action<string> onThrottledUpdate, int throttleStep = 5)
        {
            if (string.IsNullOrEmpty(delta)) return;

            string snapshot = null;
            lock (_syncRoot)
            {
                if (_isStreamFinished) return;
                _accumulator.Append(delta);
                _deltaCount++;

                if (throttleStep > 0 && _deltaCount % throttleStep == 0)
                {
                    snapshot = _accumulator.ToString();
                }
            }

            if (snapshot != null && onThrottledUpdate != null)
            {
                onThrottledUpdate(snapshot);
            }
        }

        public string FinishStream()
        {
            lock (_syncRoot)
            {
                _isStreamFinished = true;
                return _accumulator.ToString();
            }
        }

        public void Cancel()
        {
            lock (_syncRoot)
            {
                if (_cts != null)
                {
                    try { _cts.Cancel(); } catch { }
                }
            }
        }

        public void EndSending()
        {
            lock (_syncRoot)
            {
                _isSending = false;
                _isStreamFinished = true;
                if (_cts != null)
                {
                    try { _cts.Dispose(); } catch { }
                    _cts = null;
                }
            }
        }
    }
}
