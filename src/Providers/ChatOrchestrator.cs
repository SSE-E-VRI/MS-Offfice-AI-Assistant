using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using MistralOfficeAddin.Core;

namespace MistralOfficeAddin.Providers
{
    public class ChatOrchestrator : IDisposable
    {
        private readonly object _syncLock = new object();
        private IAIProvider _currentProvider;
        private CancellationTokenSource _activeStreamCts;
        private bool _isStreaming;
        private bool _disposed;

        public IAIProvider CurrentProvider
        {
            get
            {
                lock (_syncLock)
                {
                    return _currentProvider;
                }
            }
        }

        public bool IsStreaming
        {
            get
            {
                lock (_syncLock)
                {
                    return _isStreaming;
                }
            }
        }

        public ChatOrchestrator()
        {
        }

        public ChatOrchestrator(IAIProvider initialProvider)
        {
            _currentProvider = initialProvider;
        }

        /// <summary>
        /// Safely updates the active provider. If an active stream is in progress,
        /// it is cancelled before the provider is replaced and the old provider disposed.
        /// </summary>
        public void UpdateProvider(IAIProvider newProvider)
        {
            IAIProvider oldProviderToDispose = null;
            lock (_syncLock)
            {
                if (_isStreaming && _activeStreamCts != null)
                {
                    try
                    {
                        Logger.Info("ChatOrchestrator: Cancelling active stream for provider update.");
                        _activeStreamCts.Cancel();
                    }
                    catch (Exception ex)
                    {
                        Logger.Warn(string.Format("ChatOrchestrator: Error cancelling stream: {0}", ex.Message));
                    }
                }

                oldProviderToDispose = _currentProvider;
                _currentProvider = newProvider;
            }

            if (oldProviderToDispose != null && !object.ReferenceEquals(oldProviderToDispose, newProvider))
            {
                try
                {
                    oldProviderToDispose.Dispose();
                }
                catch (Exception ex)
                {
                    Logger.Warn(string.Format("ChatOrchestrator: Error disposing previous provider: {0}", ex.Message));
                }
            }
        }

        public async Task<bool> TestConnectionAsync(CancellationToken ct = default(CancellationToken))
        {
            IAIProvider provider;
            lock (_syncLock)
            {
                provider = _currentProvider;
            }

            if (provider == null)
            {
                throw new InvalidOperationException("No AI provider is currently configured.");
            }

            return await provider.TestConnectionAsync(ct).ConfigureAwait(false);
        }

        public async Task<List<AIModelInfo>> ListModelsAsync(CancellationToken ct = default(CancellationToken))
        {
            IAIProvider provider;
            lock (_syncLock)
            {
                provider = _currentProvider;
            }

            if (provider == null)
            {
                return new List<AIModelInfo>();
            }

            return await provider.ListModelsAsync(ct).ConfigureAwait(false);
        }

        public async Task<AIResponse> ChatAsync(AIRequest request, CancellationToken ct = default(CancellationToken))
        {
            IAIProvider provider;
            lock (_syncLock)
            {
                provider = _currentProvider;
            }

            if (provider == null)
            {
                throw new InvalidOperationException("No AI provider is currently configured.");
            }

            return await provider.ChatAsync(request, ct).ConfigureAwait(false);
        }

        public async Task StreamChatAsync(
            AIRequest request,
            Action<string> onDeltaReceived,
            CancellationToken ct = default(CancellationToken))
        {
            if (request == null) throw new ArgumentNullException("request");
            if (onDeltaReceived == null) throw new ArgumentNullException("onDeltaReceived");

            IAIProvider provider;
            CancellationTokenSource linkedCts;

            lock (_syncLock)
            {
                if (_disposed) throw new ObjectDisposedException("ChatOrchestrator");
                provider = _currentProvider;
                if (provider == null)
                {
                    throw new InvalidOperationException("No AI provider is currently configured.");
                }

                _activeStreamCts = new CancellationTokenSource();
                linkedCts = CancellationTokenSource.CreateLinkedTokenSource(ct, _activeStreamCts.Token);
                _isStreaming = true;
            }

            try
            {
                await provider.StreamChatAsync(request, onDeltaReceived, linkedCts.Token).ConfigureAwait(false);
            }
            finally
            {
                lock (_syncLock)
                {
                    _isStreaming = false;
                    if (_activeStreamCts != null)
                    {
                        _activeStreamCts.Dispose();
                        _activeStreamCts = null;
                    }
                }
                linkedCts.Dispose();
            }
        }

        public void CancelCurrentStream()
        {
            lock (_syncLock)
            {
                if (_isStreaming && _activeStreamCts != null)
                {
                    try
                    {
                        _activeStreamCts.Cancel();
                    }
                    catch { }
                }
            }
        }

        public bool CheckVisionSupport(string model)
        {
            IAIProvider provider;
            lock (_syncLock)
            {
                provider = _currentProvider;
            }
            if (provider == null) return false;
            return provider.CheckVisionSupport(model);
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                lock (_syncLock)
                {
                    CancelCurrentStream();
                    if (_currentProvider != null)
                    {
                        try { _currentProvider.Dispose(); } catch { }
                        _currentProvider = null;
                    }
                    _disposed = true;
                }
            }
        }
    }
}
