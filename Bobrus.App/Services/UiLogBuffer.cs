using System;
using Serilog.Events;

namespace Bobrus.App.Services;

internal static class UiLogBuffer
{
    public static event Action<LogEvent>? OnLog;

    public static void Publish(LogEvent logEvent)
    {
        OnLog?.Invoke(logEvent);
    }
}
