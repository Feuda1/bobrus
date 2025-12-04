using Serilog.Core;
using Serilog.Events;

namespace Bobrus.App.Services;

internal sealed class UiLogSink : ILogEventSink
{
    public void Emit(LogEvent logEvent)
    {
        UiLogBuffer.Publish(logEvent);
    }
}
