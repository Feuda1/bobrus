using System;
using System.Threading;
using System.Threading.Tasks;

namespace Bobrus.App.Services;
public class SetupFlowController : IDisposable
{
    private readonly CancellationTokenSource _cts = new();
    private readonly object _lock = new();
    private TaskCompletionSource _pauseTcs = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private volatile bool _isPaused;
    private bool _disposed;

    public CancellationToken Token => _cts.Token;
    public bool IsPaused => _isPaused;
    public bool IsCancellationRequested => _cts.IsCancellationRequested;

    public SetupFlowController()
    {
        _pauseTcs.TrySetResult();
    }

    public void Pause()
    {
        lock (_lock)
        {
            if (_isPaused || _disposed) return;
            _isPaused = true;
            _pauseTcs = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        }
    }

    public void Resume()
    {
        lock (_lock)
        {
            if (!_isPaused || _disposed) return;
            _isPaused = false;
            _pauseTcs.TrySetResult();
        }
    }

    public void Cancel()
    {
        lock (_lock)
        {
            if (_disposed) return;
            
            try { _cts.Cancel(); } catch { }
            if (_isPaused)
            {
                _isPaused = false;
                _pauseTcs.TrySetCanceled();
            }
        }
    }
    public async Task WaitIfPausedAsync()
    {
        Token.ThrowIfCancellationRequested();

        Task pauseTask;
        lock (_lock)
        {
            if (!_isPaused) return;
            pauseTask = _pauseTcs.Task;
        }

        try
        {
            await pauseTask.WaitAsync(Token);
        }
        catch (OperationCanceledException)
        {
            Token.ThrowIfCancellationRequested();
            throw;
        }
    }

    public void Dispose()
    {
        lock (_lock)
        {
            if (_disposed) return;
            _disposed = true;
            _pauseTcs.TrySetResult();
            
            try { _cts.Dispose(); } catch { }
        }
    }
}
