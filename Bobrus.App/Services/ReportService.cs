using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Serilog;

namespace Bobrus.App.Services;

public class ReportService
{
    private static readonly string ReportDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Bobrus", "logs");
    private readonly StringBuilder _logBuilder = new();
    private DateTime _startTime;
    private readonly List<(string Category, string Name, bool Value)> _optionsLog = new();

    public ReportService()
    {
        _startTime = DateTime.Now;
        if (!Directory.Exists(ReportDirectory)) Directory.CreateDirectory(ReportDirectory);
    }

    public void Start() 
    {
        _startTime = DateTime.Now;
        _logBuilder.Clear();
        _optionsLog.Clear();
    }

    public void LogOptions(WindowsSetupOptions options)
    {
        _optionsLog.Add(("Windows", "Меню Пуск", options.StartMenu));
        _optionsLog.Add(("Windows", "Питание USB", options.UsbPower));
        _optionsLog.Add(("Windows", "Спящий режим", options.Sleep));
        _optionsLog.Add(("Windows", "Макс. произв-ть", options.PowerPlan));
        _optionsLog.Add(("Windows", "Защитник Windows", options.Defender));
        _optionsLog.Add(("Windows", "Стандартные прил.", options.Bloat));
        _optionsLog.Add(("Windows", "Тёмная тема", options.Theme));
        _optionsLog.Add(("Windows", "UAC", options.Uac));
        _optionsLog.Add(("Windows", "Проводник", options.Explorer));
        _optionsLog.Add(("Windows", "Spec. возможности", options.Accessibility));
        _optionsLog.Add(("Windows", "Экран блокировки", options.Lock));
        _optionsLog.Add(("Windows", "Локаль (RU)", options.Locale));
        _optionsLog.Add(("Windows", "Контекстное меню", options.ContextMenu));
        _optionsLog.Add(("Windows", "Телеметрия", options.Telemetry));
        _optionsLog.Add(("Windows", "Очистка мусора", options.Cleanup));
        _optionsLog.Add(("System", "Новый пользователь", options.CreateNewUser));
        if (options.CreateNewUser) _optionsLog.Add(("System", "Пользователь", true)); 
        if (!string.IsNullOrEmpty(options.IikoServerUrl))
        {
            _optionsLog.Add(("IIKO", "Server URL", true));
            _optionsLog.Add(("IIKO", "Front", options.IikoFront));
            _optionsLog.Add(("IIKO", "Office", options.IikoOffice));
            _optionsLog.Add(("IIKO", "Chain", options.IikoChain));
            _optionsLog.Add(("IIKO", "Card", options.IikoCard));
            _optionsLog.Add(("IIKO", "Plugins", options.IikoPlugins.Any()));
        }
    }

    public void AddStep(string name, string status, string detail = "")
    {
        var time = DateTime.Now.ToString("HH:mm:ss");
        string symbol = status.ToLower() switch
        {
            "success" or "успешно" or "running" => "[+]",
            "error" or "ошибка" => "[!]",
            "warn" => "[?]",
            _ => "[*]"
        };

        if (status == "Running")
        {
             _logBuilder.AppendLine($"{time} {symbol} ЗАПУСК: {name}");
        }
        else
        {
            if (status == "Success" || status == "Успешно") symbol = "[v]";
            if (status == "Error" || status == "Ошибка") symbol = "[X]";
            
            _logBuilder.AppendLine($"{time} {symbol} {name}");
            if (!string.IsNullOrWhiteSpace(detail))
            {
                _logBuilder.AppendLine($"         > {detail}");
            }
        }
    }

    public void UpdateStep(string name, string status, string detail)
    {
        AddStep(name, status, detail);
    }

    public void LogPause(bool isPaused)
    {
        var time = DateTime.Now.ToString("HH:mm:ss");
        if (isPaused)
            _logBuilder.AppendLine($"{time} [||] --- ПАУЗА ---");
        else
            _logBuilder.AppendLine($"{time} [>]  --- ПРОДОЛЖЕНИЕ ---");
    }

    public void LogStop()
    {
        var time = DateTime.Now.ToString("HH:mm:ss");
        _logBuilder.AppendLine($"{time} [##] --- ОСТАНОВЛЕНО ПОЛЬЗОВАТЕЛЕМ ---");
    }

    public void Finish() 
    {
    }

    public string GenerateReport(string computerName)
    {
        var endTime = DateTime.Now;
        var duration = endTime - _startTime;
        var sb = new StringBuilder();

        string Border(string text, char c = '═') => new string(c, text.Length);
        string Center(string text, int width)
        {
            if (text.Length >= width) return text;
            int left = (width - text.Length) / 2;
            return new string(' ', left) + text;
        }

        sb.AppendLine("╔══════════════════════════════════════════════════════════════════════════╗");
        sb.AppendLine("║                   ОТЧЕТ ПУСКОНАЛАДКИ BOBRUS POS                          ║");
        sb.AppendLine("╚══════════════════════════════════════════════════════════════════════════╝");
        sb.AppendLine("");
        sb.AppendLine($" Дата:        {endTime:dd.MM.yyyy}");
        sb.AppendLine($" Время:       {_startTime:HH:mm:ss} - {endTime:HH:mm:ss}");
        sb.AppendLine($" Длительность: {duration:mm\\:ss}");
        sb.AppendLine($" ПК:          {computerName}");
        sb.AppendLine($" Пользователь: {Environment.UserName}");
        sb.AppendLine("");
        
        sb.AppendLine("┌──────────────────────────────────────────────────────────────────────────┐");
        sb.AppendLine("│                          ВЫБРАННЫЕ ПАРАМЕТРЫ                             │");
        sb.AppendLine("└──────────────────────────────────────────────────────────────────────────┘");
        
        var categoryGroups = _optionsLog.GroupBy(x => x.Category);
        foreach (var group in categoryGroups)
        {
            sb.AppendLine($" [{group.Key}]");
            foreach (var item in group)
            {
                var check = item.Value ? "[+]" : "[ ]";
                sb.AppendLine($"   {check} {item.Name}");
            }
            sb.AppendLine("");
        }

        sb.AppendLine("┌──────────────────────────────────────────────────────────────────────────┐");
        sb.AppendLine("│                                ЛОГ ВЫПОЛНЕНИЯ                            │");
        sb.AppendLine("└──────────────────────────────────────────────────────────────────────────┘");
        sb.AppendLine(_logBuilder.ToString());
        
        sb.AppendLine("════════════════════════════════════════════════════════════════════════════");
        sb.AppendLine(" КОНЕЦ ОТЧЕТА");
        sb.AppendLine("════════════════════════════════════════════════════════════════════════════");

        var filename = $"SetupLog_{endTime:yyyy-MM-dd_HH-mm-ss}.txt";
        var path = Path.Combine(ReportDirectory, filename);
        var reportContent = sb.ToString();
        
        File.WriteAllText(path, reportContent);

        
        return path;
    }
}
