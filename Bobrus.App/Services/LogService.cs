using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Serilog;

namespace Bobrus.App.Services;

public class LogService
{
    private readonly Action<string> _uiLogger;

    public LogService(Action<string> uiLogger)
    {
        _uiLogger = uiLogger;
        Log("=== Log Started ===", verboseOnly: true);
    }

    public void Log(string message, bool isError = false, bool verboseOnly = false)
    {
        if (isError)
        {
            Serilog.Log.Error(message);
        }
        else if (verboseOnly)
        {
            Serilog.Log.Debug(message);
        }
        else
        {
            Serilog.Log.Information(message);
        }
        if (!verboseOnly)
        {
            var uiPrefix = isError ? "✖ " : "✔ ";
            _uiLogger?.Invoke($"{uiPrefix}{message}");
        }
    }

    public void Info(string message) => Log(message);
    public void Verbose(string message) => Log(message, verboseOnly: true);
    public void Error(string message) => Log(message, isError: true);
}
