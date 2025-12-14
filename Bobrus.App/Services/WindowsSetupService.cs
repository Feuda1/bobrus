using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace Bobrus.App.Services;

public sealed record WindowsSetupOptions
{
    public bool StartMenu { get; init; }
    public bool UsbPower { get; init; }
    public bool Sleep { get; init; }
    public bool PowerPlan { get; init; }
    public bool Defender { get; init; }
    public bool SslTls { get; init; }
    public bool Bloat { get; init; }
    public bool Theme { get; init; }
    public bool Uac { get; init; }
    public bool Explorer { get; init; }
    public bool Accessibility { get; init; }
    public bool Lock { get; init; }
    public bool Locale { get; init; }
    public bool ContextMenu { get; init; }
    public bool Telemetry { get; init; }
    public bool Cleanup { get; init; }
    public bool CreateNewUser { get; init; } = false;
    public string? ComputerName { get; init; } = null;
    public bool DeleteOldUser { get; init; } = false;
    public string? IikoServerUrl { get; init; } = null;
    public bool IikoFront { get; init; } = false;
    public bool IikoOffice { get; init; } = false;
    public bool IikoChain { get; init; } = false;
    public bool IikoCard { get; init; } = true;
    public bool IikoFrontAutostart { get; init; } = true;
    public bool IikoHandCardRoll { get; init; } = true;
    public bool IikoMinimizeButton { get; init; } = true;
    public bool IikoSetServerUrl { get; init; } = true;
    public List<PluginVersion>? IikoPlugins { get; init; } = new();
    public bool SkipDefenderPrompt { get; init; } = false;
}

public sealed record WindowsSetupResult(int TotalSteps, int DoneSteps, bool Success, string Output);

internal sealed class WindowsSetupService
{
    private readonly CleaningService _cleaningService = new();
    private const string DefenderStepTitle = "Отключение защитника и брандмауэра";

    public async Task<WindowsSetupResult> ApplyAsync(WindowsSetupOptions options, IProgress<string>? progress = null, SetupFlowController? controller = null, Func<Task>? beforeIikoCallback = null)
    {
        controller ??= new SetupFlowController();
        var cancellationToken = controller.Token;

        var steps = BuildSteps(options);
        var enabledSteps = steps.Where(s => s.Enabled).ToList();
        var done = 0;
        var logBuilder = new StringBuilder();
        
        var isPhase1 = options.CreateNewUser && 
                       !string.Equals(Environment.UserName, "POS-терминал", StringComparison.OrdinalIgnoreCase);

        if (isPhase1)
        {
             var firstStep = enabledSteps.FirstOrDefault();
             if (firstStep != null && firstStep.Title.StartsWith("Создание"))
             {
                 progress?.Report($"[1/1] {firstStep.Title} (Этап 1: Создание и вход)...");
                 
                 await controller.WaitIfPausedAsync();
                 var (ok, outLog) = await RunPowerShellAsync(firstStep.Script ?? "", firstStep.Timeout);
                 
                 logBuilder.AppendLine(outLog);
                 if (!ok) 
                 {
                     progress?.Report($"⚠ Ошибка создания пользователя: {outLog}");
                     return new WindowsSetupResult(1, 0, false, logBuilder.ToString());
                 }
                 try 
                 {
                     var configPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Bobrus", "resume.json");
                     Directory.CreateDirectory(Path.GetDirectoryName(configPath)!);
                     var json = JsonSerializer.Serialize(options, new JsonSerializerOptions { WriteIndented = true });
                     await File.WriteAllTextAsync(configPath, json);

                     var exePath = Process.GetCurrentProcess().MainModule?.FileName;
                     if (exePath != null)
                     {
                          await RunPowerShellAsync(@$"Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce' -Name 'BobrusSetup' -Value '""{exePath}"" --resume'");
                     }
                 } 
                 catch (Exception ex) 
                 {
                     progress?.Report($"⚠ Ошибка сохранения конфига перезагрузки: {ex.Message}");
                 }

                 progress?.Report("⚠ Выход из системы. Войдите под новым пользователем...");
                 await Task.Delay(2000, cancellationToken);
                 Process.Start("shutdown", "/l");
                 return new WindowsSetupResult(1, 1, true, "Выход из системы...");
             }
        }

        var defenderPlanned = enabledSteps.Any(s => s.Title == DefenderStepTitle);

        if (defenderPlanned && !options.SkipDefenderPrompt)
        {
            await controller.WaitIfPausedAsync();
            await WaitForDefenderWindowClosedAsync(progress, controller);
        }

        foreach (var step in enabledSteps)
        {
            if (cancellationToken.IsCancellationRequested)
            {
                progress?.Report("⚠ Отменено пользователем.");
                break;
            }
            
            await controller.WaitIfPausedAsync();

            progress?.Report($"[{done + 1}/{enabledSteps.Count}] {step.Title}...");

            var (ok, output) = step.Executor is not null
                ? await step.Executor()
                : step.Interactive
                    ? await RunPowerShellInteractiveAsync(step.Script ?? string.Empty)
                    : await RunPowerShellAsync(step.Script ?? string.Empty, step.Timeout);

            done++;
            if (!string.IsNullOrWhiteSpace(output))
            {
                progress?.Report($"[VERBOSE] {output}");
            }

            var marker = ok ? "✔" : "✖";
            var shortOutput = SanitizeOutput(output);
            var msg = step.HideOutput || string.IsNullOrWhiteSpace(shortOutput)
                ? $"{marker} {step.Title}"
                : $"{marker} {step.Title}: {shortOutput}";
            progress?.Report(msg);
            logBuilder.AppendLine(msg);

            if (!ok && step.StopOnFail)
            {
                const string stopMessage = "Остановлено: защита не отключена. Отключите защиту от изменений (Tamper Protection) вручную: Параметры -> Обновление и безопасность -> Безопасность Windows -> Защита от вирусов и угроз -> Управление настройками -> Защита от изменений. Затем перезапустите шаг.";
                progress?.Report(stopMessage);
                logBuilder.AppendLine(stopMessage);
                break;
            }
        }

        if (enabledSteps.Count == 0)
        {
            logBuilder.AppendLine("Нет выбранных пунктов настройки Windows.");
        }
        else if (done == enabledSteps.Count)
        {
            progress?.Report("Перезапуск Explorer для применения изменений...");
            await RunPowerShellAsync("Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue; Start-Sleep -Seconds 2", TimeSpan.FromSeconds(15));
            progress?.Report("✔ Explorer перезапущен");
        }
        if (beforeIikoCallback != null)
        {
            await beforeIikoCallback();
        }
        if (!string.IsNullOrWhiteSpace(options.IikoServerUrl) && 
            (options.IikoFront || options.IikoOffice || options.IikoChain || options.IikoCard))
        {
            await controller.WaitIfPausedAsync();
            progress?.Report("");
            progress?.Report("[IIKO] Начало установки компонентов iiko...");
            
            var iikoService = new IikoSetupService();
            await iikoService.InstallIikoAsync(options, progress ?? new Progress<string>(), controller);
        }

        try 
        {
             var configPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Bobrus", "resume.json");
             if (File.Exists(configPath)) File.Delete(configPath);
        } catch { }

        return new WindowsSetupResult(enabledSteps.Count, done, done == enabledSteps.Count, logBuilder.ToString());
    }

    private static async Task<(bool Ok, string Output)> RunPowerShellAsync(string script, TimeSpan? timeout = null)
    {
        var tempLog = Path.Combine(Path.GetTempPath(), $"bobrus-win-setup-{Guid.NewGuid():N}.log");
        var escapedLog = tempLog.Replace("'", "''");

        var tempScript = Path.Combine(Path.GetTempPath(), $"bobrus-win-setup-{Guid.NewGuid():N}.ps1");
        var wrapped = new StringBuilder();
        wrapped.AppendLine("chcp 65001 | Out-Null");
        wrapped.AppendLine("$OutputEncoding = [Console]::OutputEncoding = [Text.Encoding]::UTF8");
        wrapped.AppendLine("[Threading.Thread]::CurrentThread.CurrentUICulture = 'ru-RU'");
        wrapped.AppendLine("[Threading.Thread]::CurrentThread.CurrentCulture = 'ru-RU'");
        wrapped.AppendLine("$ErrorActionPreference='Continue'");
        wrapped.AppendLine("$out = & {");
        wrapped.AppendLine(script);
        wrapped.AppendLine("} *>&1");
        wrapped.AppendLine($"$out | Out-File -FilePath '{escapedLog}' -Encoding UTF8");
        wrapped.AppendLine("$global:LASTEXITCODE = 0");
        wrapped.AppendLine("exit 0");

        await File.WriteAllTextAsync(tempScript, wrapped.ToString(), Encoding.UTF8);

        var psi = new ProcessStartInfo
        {
            FileName = "powershell",
            Arguments = $"-NoProfile -NonInteractive -ExecutionPolicy Bypass -File \"{tempScript}\"",
            UseShellExecute = true,
            Verb = "runas",
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden
        };

        using var process = Process.Start(psi);
        if (process == null)
        {
            try { if (File.Exists(tempLog)) File.Delete(tempLog); } catch { }
            try { if (File.Exists(tempScript)) File.Delete(tempScript); } catch { }
            return (false, "Не удалось запустить PowerShell");
        }

        var exited = true;
        var effectiveTimeout = timeout ?? TimeSpan.FromMinutes(2);
        using (var cts = new CancellationTokenSource(effectiveTimeout))
        {
            try
            {
                await process.WaitForExitAsync(cts.Token);
            }
            catch (OperationCanceledException)
            {
                exited = false;
            }
        }

        if (!exited && !process.HasExited)
        {
            try { process.Kill(entireProcessTree: true); } catch { }
        }

        var output = string.Empty;
        if (File.Exists(tempLog))
        {
            output = await File.ReadAllTextAsync(tempLog);
        }
        else
        {
            output = $"[C# DEBUG] Лог файл не создан: {tempLog}\n[C# DEBUG] Скрипт: {tempScript}\n[C# DEBUG] Скрипт существует: {File.Exists(tempScript)}";
        }
        
        try { if (File.Exists(tempLog)) File.Delete(tempLog); } catch { }
        try { if (File.Exists(tempScript)) File.Delete(tempScript); } catch { }
        var exitCode = process.HasExited ? process.ExitCode : -1;
        
        if (exitCode != 0 && string.IsNullOrWhiteSpace(output))
        {
            output = $"[C# DEBUG] Exit code: {exitCode}, процесс завершён: {process.HasExited}";
        }
        
        return (exitCode == 0, output);
    }

    private static async Task<(bool Ok, string Output)> RunPowerShellInteractiveAsync(string script)
    {
        var tempScript = Path.Combine(Path.GetTempPath(), $"bobrus-win-activate-{Guid.NewGuid():N}.ps1");
        var wrapped = new StringBuilder();
        wrapped.AppendLine("chcp 65001 | Out-Null");
        wrapped.AppendLine("$OutputEncoding = [Console]::OutputEncoding = [Text.Encoding]::UTF8");
        wrapped.AppendLine("[Threading.Thread]::CurrentThread.CurrentUICulture = 'ru-RU'");
        wrapped.AppendLine("[Threading.Thread]::CurrentThread.CurrentCulture = 'ru-RU'");
        wrapped.AppendLine(script);
        wrapped.AppendLine("$global:LASTEXITCODE = 0");
        await File.WriteAllTextAsync(tempScript, wrapped.ToString(), Encoding.UTF8);

        var psi = new ProcessStartInfo
        {
            FileName = "powershell",
            Arguments = $"-NoProfile -ExecutionPolicy Bypass -File \"{tempScript}\"",
            UseShellExecute = true,
            Verb = "runas",
            CreateNoWindow = false,
            WindowStyle = ProcessWindowStyle.Normal
        };

        using var process = Process.Start(psi);
        if (process == null)
        {
            return (false, "Не удалось открыть PowerShell для активации");
        }

        try
        {
            await process.WaitForExitAsync();
            return (process.ExitCode == 0, string.Empty);
        }
        finally
        {
            try { if (File.Exists(tempScript)) File.Delete(tempScript); } catch { }
        }
    }

    private static string SanitizeOutput(string output)
    {
        if (string.IsNullOrWhiteSpace(output))
            return string.Empty;

        bool IsNoise(string line)
        {
            if (line.StartsWith("At ", StringComparison.OrdinalIgnoreCase)) return true;
            if (line.StartsWith("строка", StringComparison.OrdinalIgnoreCase)) return true;
            if (line.Contains("CategoryInfo")) return true;
            if (line.Contains("FullyQualifiedErrorId")) return true;
            if (line.Contains("char:")) return true;
            return false;
        }

        var lines = output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(l => l.Trim())
            .Where(l => !string.IsNullOrWhiteSpace(l))
            .Where(l => !IsNoise(l))
            .Select(l => l.Replace("The operation completed successfully.", "Успешно"))
            .Select(l => l.Replace("Access is denied.", "Нет доступа (пропущено)"))
            .Distinct()
            .Take(6)
            .ToList();

        return string.Join(" | ", lines);
    }

    private record Step(string Title, bool Enabled, string? Script = null, Func<Task<(bool, string)>>? Executor = null, bool Interactive = false, bool StopOnFail = false, TimeSpan? Timeout = null, bool HideOutput = false);

    private IReadOnlyList<Step> BuildSteps(WindowsSetupOptions o)
    {
        return new[]
        {
            new Step("Создание нового пользователя", o.CreateNewUser && !string.IsNullOrWhiteSpace(o.ComputerName), ScriptCreateUser(o.ComputerName ?? "POS-PC", o.DeleteOldUser)),
            new Step("Настройка меню Пуск", o.StartMenu, ScriptStartMenu()),
            new Step("Отключение спящего режима", o.Sleep, ScriptSleep()),
            new Step("Отключение защитника и брандмауэра", o.Defender, ScriptDefender(), HideOutput: true),
            new Step("Настройка SSL/TLS", o.SslTls, ScriptSslTls()),
            new Step("Тёмная тема Windows", o.Theme, ScriptTheme(), Timeout: TimeSpan.FromSeconds(90)),
            new Step("Отключение заставки Windows", o.Lock, ScriptLock()),
            new Step("Отключение фоновых служб", o.Telemetry, ScriptTelemetry()),
            new Step("Отключение сна для USB", o.UsbPower, ScriptUsbPower()),
            new Step("Включить максимальную производительность", o.PowerPlan, ScriptPowerPlan()),
            new Step("Удалить стандартные приложения Windows", o.Bloat, ScriptBloat(), Timeout: TimeSpan.FromMinutes(15)),
            new Step("Отключение уведомлений администратора", o.Uac, ScriptUac()),
            new Step("Отключение залипаний клавиш", o.Accessibility, ScriptAccessibility()),
            new Step("Настройка языка и региона", o.Locale, ScriptLocale()),
            new Step("Дополнительные настройки проводника", o.ContextMenu, ScriptContextMenu()),
            new Step("Настройка проводника (Перезапуск)", o.Explorer, ScriptExplorer()), 
            new Step($"Установка iiko (плагинов: {o.IikoPlugins.Count})", o.IikoFront || o.IikoOffice || o.IikoChain || o.IikoCard, 
                Executor: () => Task.FromResult((true, "iiko installed via service"))),

            new Step("Очистка системы", o.Cleanup, Executor: async () => { await _cleaningService.RunCleanupAsync(); return (true, "Очистка завершена"); })
        };
    }

    private static string ScriptCreateUser(string computerName, bool deleteOldUsers) => $@"
$ErrorActionPreference = 'Continue'
$u = ""POS-терминал""
$computerName = ""{computerName}""

Write-Host ""Создание пользователя: $u""

try {{
    # Создание пользователя (если не существует)
    $exists = Get-LocalUser -Name $u -ErrorAction SilentlyContinue
    if (-not $exists) {{
        Write-Host ""Создание новой учетной записи...""
        net user ""$u"" /add /active:yes /passwordreq:no 2>&1 | Out-Null
        
        net localgroup Администраторы ""$u"" /add 2>&1 | Out-Null
        
        net localgroup Administrators ""$u"" /add 2>&1 | Out-Null
        
        wmic useraccount where ""Name='$u'"" set PasswordExpires=FALSE 2>&1 | Out-Null
        Write-Host ""Пользователь $u создан.""
    }} else {{
        Write-Host ""Пользователь $u уже существует.""
    }}
    
    # Удаление старых пользователей (опционально)
    if ({(deleteOldUsers ? "$true" : "$false")}) {{
        $exclude = @($u, $env:USERNAME, 'Administrator', 'Guest', 'DefaultAccount', 'WDAGUtilityAccount', 'Администратор', 'Гость')
        $usersToDelete = Get-LocalUser | Where-Object {{ $exclude -notcontains $_.Name }}
        foreach ($user in $usersToDelete) {{
            Remove-LocalUser -Name $user.Name -ErrorAction SilentlyContinue
            Write-Host ""Пользователь $($user.Name) удален.""
        }}
    }}
    
    # Переименование компьютера
    if ($env:COMPUTERNAME -ne $computerName) {{
        Write-Host ""Переименование компьютера в $computerName...""
        Rename-Computer -NewName $computerName -Force -EA SilentlyContinue
        Write-Host ""Компьютер переименован.""
    }}
    
    # Настройка автологина
    Write-Host ""Настройка автоматического входа...""
    net user ""$u"" """" *>$null
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name 'DefaultUserName' -Value $u
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name 'AutoAdminLogon' -Value '1'
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name 'DefaultDomainName' -Value $env:COMPUTERNAME
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name 'DefaultPassword' -Value """"
    
    # Отключение анимации первого входа
    New-Item -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Force -EA SilentlyContinue | Out-Null
    New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Name 'EnableFirstLogonAnimation' -Value 0 -PropertyType DWord -Force -EA SilentlyContinue | Out-Null
    
    # === УДАЛЕНИЕ АВТОЗАПУСКА У СТАРОГО ПОЛЬЗОВАТЕЛЯ (только если мы НЕ на новом пользователе) ===
    if ($env:USERNAME -ne $u) {{
        Write-Host ""Phase 1: Удаление автозапуска у старого пользователя ($env:USERNAME)...""
        try {{
            Remove-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run' -Name 'Bobrus' -Force -EA SilentlyContinue
            Remove-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run' -Name 'Rhelper' -Force -EA SilentlyContinue
            Write-Host ""Автозапуск Bobrus и Rhelper удалён у старого пользователя.""
        }} catch {{
            Write-Host ""Ошибка удаления автозапуска: $_""
        }}
    }} else {{
        Write-Host ""Phase 2: Пропуск удаления автозапуска (уже на новом пользователе).""
    }}
    
    # === АВТОЗАПУСК НАСТРАИВАЕТСЯ ТОЛЬКО В PHASE 2 (после входа в новую учётку) ===
    # Это происходит когда $env:USERNAME равен имени нового пользователя
    $rhelperExe = 'C:\Program Files (x86)\Rhelper39\RHusr_v39.exe'
    
    if ($env:USERNAME -eq $u) {{
        Write-Host ""Phase 2: Настройка для нового пользователя $u...""
        
        # 1. Сначала копируем Bobrus.exe на рабочий стол нового пользователя
        $destPath = ""C:\Users\$u\Desktop\Bobrus.exe""
        if (-not (Test-Path $destPath)) {{
            $src = Get-ChildItem -Path 'C:\Users' -Recurse -Filter 'Bobrus.exe' -EA SilentlyContinue | Where-Object {{ $_.DirectoryName -notlike ""*\$u\*"" }} | Select-Object -First 1
            if ($src) {{
                Copy-Item -Path $src.FullName -Destination $destPath -Force -EA SilentlyContinue
                Write-Host ""Bobrus.exe скопирован на рабочий стол: $destPath""
            }}
        }}
        
        # 2. Создание ярлыка Rhelper на рабочий стол
        if (Test-Path $rhelperExe) {{
            $shortcutPath = ""C:\Users\$u\Desktop\Rhelper.lnk""
            if (-not (Test-Path $shortcutPath)) {{
                $shell = New-Object -ComObject WScript.Shell
                $shortcut = $shell.CreateShortcut($shortcutPath)
                $shortcut.TargetPath = $rhelperExe
                $shortcut.WorkingDirectory = Split-Path $rhelperExe
                $shortcut.Description = ""Rhelper - Удалённый доступ""
                $shortcut.Save()
                Write-Host ""Ярлык Rhelper создан.""
            }}
        }}
        
        # 3. ТОЛЬКО ТЕПЕРЬ добавляем автозапуск (после того как файлы скопированы)
        Write-Host ""Настройка автозапуска для $u...""
        
        # Bobrus - используем путь на рабочем столе нового пользователя
        if (Test-Path $destPath) {{
            Set-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run' -Name 'Bobrus' -Value """"""$destPath"""""" -Force -EA SilentlyContinue
            Write-Host ""Bobrus добавлен в автозапуск: $destPath""
        }} else {{
            Write-Host ""ВНИМАНИЕ: Bobrus.exe не найден на рабочем столе, автозапуск не добавлен""
        }}
        
        # Rhelper
        if (Test-Path $rhelperExe) {{
            Set-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run' -Name 'Rhelper' -Value """"""$rhelperExe"""""" -Force -EA SilentlyContinue
            Write-Host ""Rhelper добавлен в автозапуск: $rhelperExe""
        }}
    }} else {{
        Write-Host ""Phase 1: Автозапуск будет настроен после входа в новую учётку.""
    }}
    
    Write-Host ""Настройка пользователя завершена.""
}} catch {{
    Write-Host ""[ERROR] Ошибка при создании пользователя: $_""
    throw
}}
";

    private async Task<(bool, string)> CleanupExecutor()
    {
        try
        {
            var results = await _cleaningService.RunCleanupAsync();
            var freed = results.Sum(r => r.BytesFreed);
            var highlights = results
                .Where(r => r.BytesFreed > 0)
                .OrderByDescending(r => r.BytesFreed)
                .Take(2)
                .Select(r => $"{r.Name}: {FormatBytes(r.BytesFreed)}")
                .ToList();

            var summary = highlights.Count > 0
                ? $"Основное: {string.Join(", ", highlights)}"
                : "Основное: нечего удалять";

            return (true, $"{summary}. Всего освобождено: {FormatBytes(freed)}");
        }
        catch (Exception ex)
        {
            return (false, ex.Message);
        }
    }

    private static string FormatBytes(long bytes)
    {
        if (bytes < 1024) return $"{bytes} Б";
        double kb = bytes / 1024.0;
        if (kb < 1024) return $"{kb:F1} КБ";
        double mb = kb / 1024.0;
        if (mb < 1024) return $"{mb:F1} МБ";
        double gb = mb / 1024.0;
        return $"{gb:F1} ГБ";
    }

    private static string ScriptStartMenu() => @"
$ErrorActionPreference='SilentlyContinue'

# ===== ПАНЕЛЬ ЗАДАЧ =====

# Скрыть поиск с панели задач
Try { reg add 'HKCU\Software\Microsoft\Windows\CurrentVersion\Search' /v SearchboxTaskbarMode /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKCU\Software\Microsoft\Windows\CurrentVersion\Search' /v SearchMode /t REG_DWORD /d 0 /f | Out-Null } Catch {}

# УБРАТЬ КНОПКУ ПРОСМОТРА ЗАДАЧ (Task View)
Try { reg add 'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' /v ShowTaskViewButton /t REG_DWORD /d 0 /f | Out-Null } Catch {}

# Отключить виджеты на панели задач (Windows 11)
Try { reg add 'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' /v TaskbarDa /t REG_DWORD /d 0 /f | Out-Null } Catch {}

# Отключить Chat/Teams на панели задач (Windows 11)
Try { reg add 'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' /v TaskbarMn /t REG_DWORD /d 0 /f | Out-Null } Catch {}

# Отключить Copilot на панели задач (Windows 11)
Try { reg add 'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' /v ShowCopilotButton /t REG_DWORD /d 0 /f | Out-Null } Catch {}

# Отключить новости и интересы (виджеты Windows 10)
Try { reg add 'HKCU\Software\Microsoft\Windows\CurrentVersion\Feeds' /v ShellFeedsTaskbarViewMode /t REG_DWORD /d 2 /f | Out-Null } Catch {}

# Отключить уведомления OneDrive в проводнике
Try { reg add 'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' /v ShowSyncProviderNotifications /t REG_DWORD /d 0 /f | Out-Null } Catch {}

# ===== МЕНЮ ПУСК - ПОЛНАЯ ОЧИСТКА =====

# Отключить рекомендации в Пуске (Windows 11)
Try { reg add 'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' /v Start_IrisRecommendations /t REG_DWORD /d 0 /f | Out-Null } Catch {}

# Отключить рекламу и предложения в Пуске
Try { reg add 'HKCU\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' /v SystemPaneSuggestionsEnabled /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKCU\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' /v SubscribedContent-338388Enabled /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKCU\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' /v SubscribedContent-338389Enabled /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKCU\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' /v SubscribedContent-338393Enabled /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKCU\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' /v SilentInstalledAppsEnabled /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKCU\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' /v SoftLandingEnabled /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKCU\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' /v PreInstalledAppsEnabled /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKCU\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' /v OemPreInstalledAppsEnabled /t REG_DWORD /d 0 /f | Out-Null } Catch {}

# Отключить ""Показывать недавно добавленные приложения""
Try { reg add 'HKCU\Software\Microsoft\Windows\CurrentVersion\Start' /v ShowRecentList /t REG_DWORD /d 0 /f | Out-Null } Catch {}

# Отключить ""Показывать часто используемые приложения""
Try { reg add 'HKCU\Software\Microsoft\Windows\CurrentVersion\Start' /v ShowFrequentList /t REG_DWORD /d 0 /f | Out-Null } Catch {}

# ===== WINDOWS 11 - ОЧИСТКА ЗАКРЕПЛЁННЫХ ПРИЛОЖЕНИЙ =====
Try {
    $pinnedPath = ""$env:LOCALAPPDATA\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState""
    if (Test-Path $pinnedPath) {
        # Останавливаем процесс StartMenuExperienceHost
        Stop-Process -Name StartMenuExperienceHost -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 500

        # Удаляем файлы с закреплёнными приложениями
        Remove-Item -Path ""$pinnedPath\start*.bin"" -Force -ErrorAction SilentlyContinue
        Remove-Item -Path ""$pinnedPath\start2.bin"" -Force -ErrorAction SilentlyContinue

        # Создаём пустой файл start2.bin (JSON с пустым списком)
        # Для Windows 11 формат бинарный, но можно попробовать JSON
        $emptyPinned = '{""pinnedList"":[]}'
        [System.IO.File]::WriteAllText(""$pinnedPath\start2.bin"", $emptyPinned)
    }
} Catch {}

# ===== WINDOWS 10 - ОЧИСТКА ПЛИТОК =====
Try {
    # Создаём пустой layout для Windows 10
    $layoutPath = ""$env:LOCALAPPDATA\Microsoft\Windows\Shell\LayoutModification.xml""
    $emptyLayout = @'
<?xml version=""1.0"" encoding=""utf-8""?>
<LayoutModificationTemplate xmlns:defaultlayout=""http://schemas.microsoft.com/Start/2014/FullDefaultLayout""
    xmlns:start=""http://schemas.microsoft.com/Start/2014/StartLayout""
    xmlns=""http://schemas.microsoft.com/Start/2014/LayoutModification""
    Version=""1"">
  <LayoutOptions StartTileGroupCellWidth=""6"" />
  <DefaultLayoutOverride>
    <StartLayoutCollection>
      <defaultlayout:StartLayout GroupCellWidth=""6"" />
    </StartLayoutCollection>
  </DefaultLayoutOverride>
</LayoutModificationTemplate>
'@
    $emptyLayout | Out-File -FilePath $layoutPath -Encoding UTF8 -Force
} Catch {}

# Попытка очистить плитки через реестр (экспорт/импорт)
Try {
    # Путь к данным о плитках в реестре
    $tilesRegPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\CloudStore\Store\Cache\DefaultAccount'
    if (Test-Path $tilesRegPath) {
        # Находим и удаляем ключи с плитками
        Get-ChildItem -Path $tilesRegPath -Recurse -ErrorAction SilentlyContinue | Where-Object {
            $_.Name -match 'start\.tilegrid'
        } | ForEach-Object {
            Remove-Item -Path $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
} Catch {}

# Удаление ярлыков bloatware из меню Пуск
Try {
    $startMenuPaths = @(
        ""$env:APPDATA\Microsoft\Windows\Start Menu\Programs"",
        ""$env:ProgramData\Microsoft\Windows\Start Menu\Programs""
    )
    $bloatwarePatterns = @('Xbox', 'Microsoft Store', 'Mail', 'Calendar', 'Cortana', 'Groove', 'Movies', 'News', 'Weather', 'OneNote', 'Skype', 'Teams', 'Your Phone', 'Mixed Reality', 'Paint 3D', '3D Viewer')

    foreach ($path in $startMenuPaths) {
        if (Test-Path $path) {
            Get-ChildItem -Path $path -Recurse -Include '*.lnk' -ErrorAction SilentlyContinue | Where-Object {
                $name = $_.BaseName
                $bloatwarePatterns | Where-Object { $name -match $_ }
            } | Remove-Item -Force -ErrorAction SilentlyContinue
        }
    }
} Catch {}

# Explorer будет перезапущен в конце всех операций

Write-Output 'Пуск очищен.'

$global:LASTEXITCODE = 0
";

    private static string ScriptUsbPower() => string.Join("; ",
        "Try { powercfg /SETACVALUEINDEX SCHEME_CURRENT SUB_USB a2443a6e-c51b-4619-8126-92e6f8b02f75 0 *> $null } Catch {}",
        "Try { powercfg /SETDCVALUEINDEX SCHEME_CURRENT SUB_USB a2443a6e-c51b-4619-8126-92e6f8b02f75 0 *> $null } Catch {}",
        "Try { powercfg -SetActive SCHEME_CURRENT *> $null } Catch {}",
        "$global:LASTEXITCODE = 0");

    private static string ScriptSleep() => string.Join("; ",
        "powercfg -hibernate off",
        "powercfg /x -standby-timeout-ac 0",
        "powercfg /x -standby-timeout-dc 0",
        "powercfg /x -hibernate-timeout-ac 0",
        "powercfg /x -hibernate-timeout-dc 0",
        "$global:LASTEXITCODE = 0");

    private static string ScriptPowerPlan() => @"
$ErrorActionPreference='SilentlyContinue'

# GUIDs планов питания
$ultimateGuid = 'e9a42b02-d5df-448d-aa00-03f14749eb61'  # Ultimate Performance (скрытый)
$highPerfGuid = '8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c'  # High Performance

$activated = $false

# Попытка 1: Ultimate Performance (может быть скрыт)
Try {
    # Сначала пробуем создать копию Ultimate Performance (разблокирует его)
    $result = powercfg -duplicatescheme $ultimateGuid 2>&1
    if ($LASTEXITCODE -eq 0) {
        powercfg -setactive $ultimateGuid *> $null
        if ($LASTEXITCODE -eq 0) {
            $activated = $true
            Write-Output 'Ultimate Performance активирован'
        }
    }
} Catch {}

# Попытка 2: High Performance (fallback)
if (-not $activated) {
    Try {
        powercfg -setactive $highPerfGuid *> $null
        if ($LASTEXITCODE -eq 0) {
            $activated = $true
            Write-Output 'High Performance активирован'
        }
    } Catch {}
}

# Попытка 3: Если ни один не сработал, ищем доступные планы
if (-not $activated) {
    Try {
        $plans = powercfg -list | Select-String -Pattern 'GUID:\s*(\S+).*\((.*)\)'
        foreach ($plan in $plans) {
            if ($plan -match 'High|Высок|Ultimate|Максим') {
                $guid = $plan.Matches.Groups[1].Value
                powercfg -setactive $guid *> $null
                if ($LASTEXITCODE -eq 0) {
                    $activated = $true
                    Write-Output ""Активирован план: $($plan.Matches.Groups[2].Value)""
                    break
                }
            }
        }
    } Catch {}
}

# Настройка параметров текущего плана для максимальной производительности
Try {
    # Минимальное состояние процессора = 5% (чтобы не держать 100% постоянно)
    powercfg /SETACVALUEINDEX SCHEME_CURRENT SUB_PROCESSOR PROCTHROTTLEMIN 5 *> $null
    powercfg /SETDCVALUEINDEX SCHEME_CURRENT SUB_PROCESSOR PROCTHROTTLEMIN 5 *> $null

    # Максимальное состояние процессора = 100%
    powercfg /SETACVALUEINDEX SCHEME_CURRENT SUB_PROCESSOR PROCTHROTTLEMAX 100 *> $null
    powercfg /SETDCVALUEINDEX SCHEME_CURRENT SUB_PROCESSOR PROCTHROTTLEMAX 100 *> $null

    # Разрешить простой процессора
    powercfg /SETACVALUEINDEX SCHEME_CURRENT SUB_PROCESSOR IDLEDISABLE 0 *> $null
    powercfg /SETDCVALUEINDEX SCHEME_CURRENT SUB_PROCESSOR IDLEDISABLE 0 *> $null

    # Политика охлаждения = активная
    powercfg /SETACVALUEINDEX SCHEME_CURRENT SUB_PROCESSOR SYSCOOLPOL 1 *> $null
    powercfg /SETDCVALUEINDEX SCHEME_CURRENT SUB_PROCESSOR SYSCOOLPOL 1 *> $null

    # Применить изменения
    powercfg -SetActive SCHEME_CURRENT *> $null
} Catch {}

# Отключение быстрого запуска (Fast Startup) - предотвращает проблемы с аптаймом
Try { reg add ""HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Power"" /v HiberbootEnabled /t REG_DWORD /d 0 /f | Out-Null } Catch {}

$global:LASTEXITCODE = 0
";

    private static string ScriptDefender() => @"
$ErrorActionPreference='SilentlyContinue'

# ===== ГРУБОЕ ОТКЛЮЧЕНИЕ DEFENDER И БРАНДМАУЭРА =====

# Реестр - Defender (политики)
reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows Defender' /v DisableAntiSpyware /t REG_DWORD /d 1 /f *>$null
reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows Defender' /v DisableAntiVirus /t REG_DWORD /d 1 /f *>$null
reg add 'HKLM\SOFTWARE\Microsoft\Windows Defender' /v DisableAntiSpyware /t REG_DWORD /d 1 /f *>$null
reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection' /v DisableRealtimeMonitoring /t REG_DWORD /d 1 /f *>$null
reg add 'HKLM\SOFTWARE\Microsoft\Windows Defender\Real-Time Protection' /v DisableRealtimeMonitoring /t REG_DWORD /d 1 /f *>$null

# SmartScreen
reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\System' /v EnableSmartScreen /t REG_DWORD /d 0 /f *>$null
reg add 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer' /v SmartScreenEnabled /t REG_SZ /d 'Off' /f *>$null

# Скрыть Security Center
reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows Defender Security Center\Notifications' /v DisableNotifications /t REG_DWORD /d 1 /f *>$null
reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows Defender Security Center\Systray' /v HideSystray /t REG_DWORD /d 1 /f *>$null
reg delete 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run' /v SecurityHealth /f *>$null

# Службы Defender - отключить через реестр
reg add 'HKLM\SYSTEM\CurrentControlSet\Services\WinDefend' /v Start /t REG_DWORD /d 4 /f *>$null
reg add 'HKLM\SYSTEM\CurrentControlSet\Services\WdNisSvc' /v Start /t REG_DWORD /d 4 /f *>$null
reg add 'HKLM\SYSTEM\CurrentControlSet\Services\SecurityHealthService' /v Start /t REG_DWORD /d 4 /f *>$null

# ===== БРАНДМАУЭР - ПОЛНОЕ ОТКЛЮЧЕНИЕ =====

# Через netsh (самый надёжный способ)
netsh advfirewall set allprofiles state off *>$null

# Через PowerShell
Try { Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled False -ErrorAction SilentlyContinue } Catch {}

# Служба брандмауэра - отключить
sc.exe config MpsSvc start= disabled *>$null
net stop MpsSvc /y *>$null
reg add 'HKLM\SYSTEM\CurrentControlSet\Services\MpsSvc' /v Start /t REG_DWORD /d 4 /f *>$null

# Реестр брандмауэра
reg add 'HKLM\SOFTWARE\Policies\Microsoft\WindowsFirewall\DomainProfile' /v EnableFirewall /t REG_DWORD /d 0 /f *>$null
reg add 'HKLM\SOFTWARE\Policies\Microsoft\WindowsFirewall\StandardProfile' /v EnableFirewall /t REG_DWORD /d 0 /f *>$null
reg add 'HKLM\SOFTWARE\Policies\Microsoft\WindowsFirewall\PublicProfile' /v EnableFirewall /t REG_DWORD /d 0 /f *>$null
reg add 'HKLM\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\DomainProfile' /v EnableFirewall /t REG_DWORD /d 0 /f *>$null
reg add 'HKLM\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile' /v EnableFirewall /t REG_DWORD /d 0 /f *>$null
reg add 'HKLM\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\PublicProfile' /v EnableFirewall /t REG_DWORD /d 0 /f *>$null

# PowerShell команды (Set-MpPreference) для надежности
Try { Set-MpPreference -DisableRealtimeMonitoring $true -ErrorAction SilentlyContinue } Catch {}
Try { Set-MpPreference -DisableBehaviorMonitoring $true -ErrorAction SilentlyContinue } Catch {}
Try { Set-MpPreference -DisableBlockAtFirstSeen $true -ErrorAction SilentlyContinue } Catch {}
Try { Set-MpPreference -DisableIOAVProtection $true -ErrorAction SilentlyContinue } Catch {}
Try { Set-MpPreference -DisablePrivacyMode $true -ErrorAction SilentlyContinue } Catch {}
Try { Set-MpPreference -SignatureDisableUpdateOnStartupWithoutEngine $true -ErrorAction SilentlyContinue } Catch {}
Try { Set-MpPreference -DisableArchiveScanning $true -ErrorAction SilentlyContinue } Catch {}
Try { Set-MpPreference -DisableIntrusionPreventionSystem $true -ErrorAction SilentlyContinue } Catch {}
Try { Set-MpPreference -DisableScriptScanning $true -ErrorAction SilentlyContinue } Catch {}

Try { sc.exe stop WinDefend | Out-Null } Catch {}
Try { sc.exe config WinDefend start= disabled | Out-Null } Catch {}

Write-Output 'Defender и брандмауэр отключены'
$global:LASTEXITCODE = 0
";

    private static string ScriptBloat() => @"
$ErrorActionPreference='SilentlyContinue'

# Полный список bloatware для POS терминала
$apps = @(
    # Xbox и игры
    'Microsoft.XboxApp',
    'Microsoft.XboxGamingOverlay',
    'Microsoft.Xbox.TCUI',
    'Microsoft.XboxGameCallableUI',
    'Microsoft.XboxIdentityProvider',
    'Microsoft.XboxSpeechToTextOverlay',
    'Microsoft.GamingApp',
    'Microsoft.GamingServices',

    # Мультимедиа
    'Microsoft.ZuneMusic',
    'Microsoft.ZuneVideo',
    'Microsoft.WindowsMediaPlayer',
    'Microsoft.Media.PlayReadyClient',

    # Помощь и справка
    'Microsoft.GetHelp',
    'Microsoft.Getstarted',
    'Microsoft.WindowsTips',

    # Новости и погода
    'Microsoft.BingNews',
    'Microsoft.BingWeather',
    'Microsoft.BingFinance',
    'Microsoft.BingSports',
    'Microsoft.BingTravel',
    'Microsoft.BingHealthAndFitness',
    'Microsoft.BingFoodAndDrink',

    # Microsoft приложения
    'Microsoft.WindowsFeedbackHub',
    'Microsoft.WindowsCommunicationsApps',
    'Microsoft.People',
    'Microsoft.MicrosoftSolitaireCollection',
    'Microsoft.MicrosoftOfficeHub',
    'Microsoft.Office.OneNote',
    'Microsoft.Microsoft3DViewer',
    'Microsoft.3DBuilder',
    'Microsoft.Print3D',
    'Microsoft.MixedReality.Portal',
    'Microsoft.WindowsMaps',
    'Microsoft.Todos',
    'Microsoft.Whiteboard',
    'Microsoft.PowerAutomateDesktop',
    'Microsoft.MicrosoftJournal',

    # Teams и коммуникации
    'MicrosoftTeams',
    'Microsoft.Teams',
    'Microsoft.SkypeApp',
    'Microsoft.Messaging',
    'Microsoft.YourPhone',
    'Microsoft.WindowsPhone',

    # Edge и Copilot
    'Microsoft.MicrosoftEdge.Stable',
    'Microsoft.MicrosoftEdge',
    'Microsoft.MicrosoftEdgeDevToolsClient',
    'Microsoft.Copilot',
    'Microsoft.Windows.Ai.Copilot.Provider',

    # Cortana и виджеты
    'Microsoft.549981C3F5F10', # Cortana
    'Microsoft.Windows.Cortana',
    'MicrosoftWindows.Client.WebExperience', # Виджеты

    # Clipchamp и редакторы
    'Clipchamp.Clipchamp',
    'Microsoft.Windows.Photos', # Фото (если не нужно)
    'Microsoft.Paint',
    'Microsoft.Paint3D',
    'Microsoft.MSPaint',

    # OneDrive
    'Microsoft.OneDrive',
    'Microsoft.OneDriveSync',
    'Microsoft.SkyDrive',

    # Камера и другое
    'Microsoft.WindowsCamera',
    'Microsoft.WindowsSoundRecorder',
    'Microsoft.WindowsAlarms',
    'Microsoft.WindowsCalculator', # Оставить если нужен калькулятор
    'Microsoft.ScreenSketch',
    'Microsoft.Windows.DevHome',

    # Магазин (опционально - может понадобиться для обновлений)
    # 'Microsoft.WindowsStore',
    # 'Microsoft.StorePurchaseApp',

    # Другие ненужные
    'Microsoft.Advertising.Xaml',
    'Microsoft.Services.Store.Engagement',
    'Microsoft.AppConnector',
    'Microsoft.ConnectivityStore',
    'Microsoft.CommsPhone',
    'Microsoft.Wallet',
    'Microsoft.PPIProjection',
    'Microsoft.WebMediaExtensions',
    'Microsoft.VP9VideoExtensions',
    'Microsoft.HEIFImageExtension',
    'Microsoft.WebpImageExtension',
    'Microsoft.RawImageExtension',
    'Microsoft.HEVCVideoExtension',
    'Microsoft.LanguageExperiencePackru-RU',
    'Microsoft.MixedReality.Portal',
    'Microsoft.Holographic.FirstRun',

    # LinkedIn
    'Microsoft.LinkedIn',

    # Sway
    'Microsoft.Office.Sway',

    # Network Speed Test
    'Microsoft.NetworkSpeedTest',

    # Remote Desktop (если не нужен)
    # 'Microsoft.RemoteDesktop',

    # Sticky Notes (если не нужен)
    'Microsoft.MicrosoftStickyNotes',

    # Quick Assist
    'MicrosoftCorporationII.QuickAssist',

    # Family
    'MicrosoftCorporationII.MicrosoftFamily',

    # Outlook (новый)
    'Microsoft.OutlookForWindows',

    # Power Automate
    'Microsoft.PowerAutomateDesktop'
)

# Удаление установленных пакетов
foreach ($app in $apps) {
    Try { Get-AppxPackage -Name $app -AllUsers | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue } Catch {}
    Try { Get-AppxPackage -Name $app | Remove-AppxPackage -ErrorAction SilentlyContinue } Catch {}
}

# Удаление provisioned пакетов (чтобы не устанавливались для новых пользователей)
Try {
    $provisioned = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
    foreach ($app in $apps) {
        $provisioned | Where-Object { $_.DisplayName -like ""*$app*"" } | ForEach-Object {
            Try { Remove-AppxProvisionedPackage -Online -PackageName $_.PackageName -ErrorAction SilentlyContinue | Out-Null } Catch {}
        }
    }
} Catch {}

# Отключение Consumer Features (реклама в Пуске)
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\CloudContent' /v DisableWindowsConsumerFeatures /t REG_DWORD /d 1 /f | Out-Null } Catch {}
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\CloudContent' /v DisableSoftLanding /t REG_DWORD /d 1 /f | Out-Null } Catch {}
Try { reg add 'HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' /v SilentInstalledAppsEnabled /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' /v SystemPaneSuggestionsEnabled /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' /v SoftLandingEnabled /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' /v RotatingLockScreenEnabled /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' /v RotatingLockScreenOverlayEnabled /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' /v SubscribedContent-310093Enabled /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' /v SubscribedContent-338388Enabled /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' /v SubscribedContent-338389Enabled /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' /v SubscribedContent-338393Enabled /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' /v SubscribedContent-353694Enabled /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' /v SubscribedContent-353696Enabled /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' /v ContentDeliveryAllowed /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' /v OemPreInstalledAppsEnabled /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' /v PreInstalledAppsEnabled /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' /v PreInstalledAppsEverEnabled /t REG_DWORD /d 0 /f | Out-Null } Catch {}

# Удаление OneDrive
Try { taskkill /F /IM OneDrive.exe *> $null } Catch {}
Try {
    if (Test-Path ""$env:SystemRoot\System32\OneDriveSetup.exe"") {
        & ""$env:SystemRoot\System32\OneDriveSetup.exe"" /uninstall *> $null
    }
} Catch {}
Try {
    if (Test-Path ""$env:SystemRoot\SysWOW64\OneDriveSetup.exe"") {
        & ""$env:SystemRoot\SysWOW64\OneDriveSetup.exe"" /uninstall *> $null
    }
} Catch {}

# Отключение Cortana
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search' /v AllowCortana /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search' /v AllowCortanaAboveLock /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search' /v DisableWebSearch /t REG_DWORD /d 1 /f | Out-Null } Catch {}
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search' /v ConnectedSearchUseWeb /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKCU\SOFTWARE\Policies\Microsoft\Windows\Explorer' /v DisableSearchBoxSuggestions /t REG_DWORD /d 1 /f | Out-Null } Catch {}
Try { reg add 'HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search' /v BingSearchEnabled /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search' /v DisableSearchBoxSuggestions /t REG_DWORD /d 1 /f | Out-Null } Catch {}
Try { reg add 'HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search' /v IsAADCloudSearchEnabled /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search' /v DeviceHistoryEnabled /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search' /v CortanaConsent /t REG_DWORD /d 0 /f | Out-Null } Catch {}

# Отключение виджетов
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Dsh' /v AllowNewsAndInterests /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced' /v TaskbarDa /t REG_DWORD /d 0 /f | Out-Null } Catch {}

# Отключение всех автозагрузок кроме Bobrus
Try {
    # Реестр HKCU Run
    $runKey = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
    Get-ItemProperty -Path $runKey -ErrorAction SilentlyContinue | 
        ForEach-Object { $_.PSObject.Properties } | 
        Where-Object { $_.Name -notlike 'PS*' -and $_.Name -notlike 'Bobrus*' } |
        ForEach-Object { Remove-ItemProperty -Path $runKey -Name $_.Name -Force -ErrorAction SilentlyContinue }
} Catch {}

Try {
    # Реестр HKLM Run
    $runKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
    Get-ItemProperty -Path $runKey -ErrorAction SilentlyContinue | 
        ForEach-Object { $_.PSObject.Properties } | 
        Where-Object { $_.Name -notlike 'PS*' -and $_.Name -notlike 'Bobrus*' -and $_.Name -notlike 'SecurityHealth*' } |
        ForEach-Object { Remove-ItemProperty -Path $runKey -Name $_.Name -Force -ErrorAction SilentlyContinue }
} Catch {}

Try {
    # Папка автозагрузки пользователя
    $startupPath = [Environment]::GetFolderPath('Startup')
    Get-ChildItem -Path $startupPath -ErrorAction SilentlyContinue | 
        Where-Object { $_.Name -notlike '*Bobrus*' } |
        ForEach-Object { Remove-Item -Path $_.FullName -Force -ErrorAction SilentlyContinue }
} Catch {}

Try {
    # Общая папка автозагрузки
    $commonStartup = [Environment]::GetFolderPath('CommonStartup')
    Get-ChildItem -Path $commonStartup -ErrorAction SilentlyContinue | 
        Where-Object { $_.Name -notlike '*Bobrus*' } |
        ForEach-Object { Remove-Item -Path $_.FullName -Force -ErrorAction SilentlyContinue }
} Catch {}

$global:LASTEXITCODE = 0
";

    private static string ScriptTheme() => @"
$ErrorActionPreference='Continue'

Write-Host ""Применение тёмной темы...""

# Тёмная тема
reg add 'HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' /v AppsUseLightTheme /t REG_DWORD /d 0 /f
reg add 'HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' /v SystemUsesLightTheme /t REG_DWORD /d 0 /f
reg add 'HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' /v EnableTransparency /t REG_DWORD /d 0 /f

# Отключение системных звуков (звуковая схема Без звука)
Try { reg add ""HKCU\AppEvents\Schemes"" /ve /t REG_SZ /d "".None"" /f | Out-Null } Catch {}

# Создаём изображение с серым фоном (цвет Ураган #3A3A3A)
# Сохраняем в постоянную папку Bobrus, а не во временную
$bobrusData = Join-Path $env:ProgramData 'Bobrus'
if (!(Test-Path $bobrusData)) { New-Item -Path $bobrusData -ItemType Directory -Force | Out-Null }
$wallpaperPath = Join-Path $bobrusData 'gray-wallpaper.bmp'

# URL изображения с серым фоном (1920x1080 сплошной серый цвет)
$imageUrl = 'https:

Try {
    # Скачиваем изображение
    Invoke-WebRequest -Uri $imageUrl -OutFile $wallpaperPath -UseBasicParsing -TimeoutSec 10
} Catch {
    # Если не удалось скачать, создаём локально через Add-Type
    Write-Host ""Загрузка фона не удалась, генерируем стандартный серый фон...""
    Try {
        Add-Type -AssemblyName System.Drawing
        $bmp = New-Object System.Drawing.Bitmap(1920, 1080)
        $graphics = [System.Drawing.Graphics]::FromImage($bmp)
        $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(58, 58, 58))
        $graphics.FillRectangle($brush, 0, 0, 1920, 1080)
        $bmp.Save($wallpaperPath, [System.Drawing.Imaging.ImageFormat]::Bmp)
        $graphics.Dispose()
        $brush.Dispose()
        $bmp.Dispose()
        Write-Host ""Фон создан.""
    } Catch {
        Write-Host ""Ошибка генерации фона: $_""
        # В крайнем случае используем сплошной цвет через реестр
        reg add 'HKCU\Control Panel\Colors' /v Background /t REG_SZ /d '58 58 58' /f
    }
}

# Очистка кеша обоев
$transcodedPath = Join-Path $env:APPDATA 'Microsoft\Windows\Themes\TranscodedWallpaper'
$cachedPath = Join-Path $env:APPDATA 'Microsoft\Windows\Themes\CachedFiles'
if (Test-Path $transcodedPath) { Remove-Item $transcodedPath -Force }
if (Test-Path $cachedPath) { Remove-Item $cachedPath -Recurse -Force }

# Устанавливаем обои через реестр
reg add 'HKCU\Control Panel\Desktop' /v Wallpaper /t REG_SZ /d ""$wallpaperPath"" /f
reg add 'HKCU\Control Panel\Desktop' /v WallpaperStyle /t REG_SZ /d 10 /f
reg add 'HKCU\Control Panel\Desktop' /v TileWallpaper /t REG_SZ /d 0 /f
reg add 'HKCU\Control Panel\Colors' /v Background /t REG_SZ /d '58 58 58' /f

# Применяем через Windows API (более надёжный метод)
Try {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class BobrusWallpaper {
    [DllImport(""user32.dll"", CharSet = CharSet.Auto)]
    public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
'@ 
} Catch {
    # Попробуем продолжить, если тип уже был загружен ранее
}

Try {
    [BobrusWallpaper]::SystemParametersInfo(0x0014, 0, $wallpaperPath, 0x0001 -bor 0x0002)
} Catch {
    Write-Host ""Ошибка API применения обоев: $_""
}

# Дополнительно применяем изменения через rundll32
rundll32.exe user32.dll,UpdatePerUserSystemParameters 1 True
Start-Sleep -Seconds 1

# Explorer будет перезапущен в конце пусконаладки

Write-Output 'Тёмная тема установлена'
$global:LASTEXITCODE = 0
";

    private static string ScriptUac() => string.Join("; ",
        "reg add \"HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\System\" /v ConsentPromptBehaviorAdmin /t REG_DWORD /d 0 /f | Out-Null",
        "reg add \"HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\System\" /v PromptOnSecureDesktop /t REG_DWORD /d 0 /f | Out-Null");

    private static string ScriptExplorer() => @"
$ErrorActionPreference='SilentlyContinue'

function Set-RegDword {
    param($Path, $Name, $Val)
    Try {
        if (!(Test-Path $Path)) { New-Item -Path $Path -Force -ErrorAction SilentlyContinue | Out-Null }
        New-ItemProperty -Path $Path -Name $Name -Value $Val -PropertyType DWord -Force -ErrorAction SilentlyContinue | Out-Null
    } Catch {}
}

# Скрыть расширения файлов (0 = show? No, HideFileExt=0 means SHOW)
Set-RegDword 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' 'HideFileExt' 0
Set-RegDword 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' 'LaunchTo' 1

# Отображение значков на рабочем столе
# {20D04...} = Этот компьютер
# {59031...} = Файлы пользователя
Set-RegDword 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel' '{20D04FE0-3AEA-1069-A2D8-08002B30309D}' 0
Set-RegDword 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel' '{59031a47-3f72-44a7-89c5-5595fe6b30ee}' 1

$global:LASTEXITCODE = 0
";

    private static string ScriptAccessibility() => string.Join("; ",
        "Try { reg add \"HKCU\\Control Panel\\Accessibility\\StickyKeys\" /v Flags /t REG_SZ /d 506 /f | Out-Null } Catch {}",
        "Try { reg add \"HKCU\\Control Panel\\Accessibility\\Keyboard Response\" /v Flags /t REG_SZ /d 122 /f | Out-Null } Catch {}",
        "Try { reg add \"HKCU\\Control Panel\\Accessibility\\ToggleKeys\" /v Flags /t REG_SZ /d 58 /f | Out-Null } Catch {}",
        "$global:LASTEXITCODE = 0");

    private static string ScriptLock() => string.Join("; ",
        "Try { reg add \"HKCU\\Control Panel\\Desktop\" /v ScreenSaveActive /t REG_SZ /d 0 /f | Out-Null } Catch {}",
        "Try { reg add \"HKCU\\Control Panel\\Desktop\" /v ScreenSaveTimeOut /t REG_SZ /d 0 /f | Out-Null } Catch {}",
        "Try { reg add \"HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Personalization\" /v NoLockScreen /t REG_DWORD /d 1 /f | Out-Null } Catch {}",
        "$global:LASTEXITCODE = 0");

    private static string ScriptLocale() => @"
$ErrorActionPreference='SilentlyContinue'

# Установить регион Россия
Try { Set-WinHomeLocation -GeoId 203 } Catch {}
Try { Set-Culture ru-RU } Catch {}
Try { Set-WinSystemLocale ru-RU } Catch {}

# Добавить русский И английский язык клавиатуры
Try {
    $langList = New-WinUserLanguageList -Language 'ru-RU'
    $langList.Add('en-US')
    Set-WinUserLanguageList -LanguageList $langList -Force
} Catch {}

# Альтернативный способ через реестр (для надежности)
Try {
    # Preload - список раскладок клавиатуры
    # 00000419 = Русская, 00000409 = Английская (США)
    reg add 'HKCU\Keyboard Layout\Preload' /v 1 /t REG_SZ /d '00000419' /f | Out-Null
    reg add 'HKCU\Keyboard Layout\Preload' /v 2 /t REG_SZ /d '00000409' /f | Out-Null
} Catch {}

    # Исправление кодировок (кракозябры)
    # Форсируем использование таблицы 1251 (кириллица) вместо 1250/1252
    Try {
        reg add 'HKLM\SYSTEM\CurrentControlSet\Control\Nls\CodePage' /v 1250 /t REG_SZ /d 'c_1251.nls' /f | Out-Null
        reg add 'HKLM\SYSTEM\CurrentControlSet\Control\Nls\CodePage' /v 1251 /t REG_SZ /d 'c_1251.nls' /f | Out-Null
        reg add 'HKLM\SYSTEM\CurrentControlSet\Control\Nls\CodePage' /v 1252 /t REG_SZ /d 'c_1251.nls' /f | Out-Null
        
        # Font Substitutes (дополнительная мера)
        reg add 'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\FontSubstitutes' /v 'Segoe UI,0' /t REG_SZ /d 'Segoe UI,204' /f | Out-Null
        reg add 'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\FontSubstitutes' /v 'Tahoma,0' /t REG_SZ /d 'Tahoma,204' /f | Out-Null
        reg add 'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\FontSubstitutes' /v 'Arial,0' /t REG_SZ /d 'Arial,204' /f | Out-Null
    } Catch {}

$global:LASTEXITCODE = 0
";

    private static string ScriptContextMenu() => @"
$ErrorActionPreference='SilentlyContinue'

# ===== КЛАССИЧЕСКОЕ КОНТЕКСТНОЕ МЕНЮ (Windows 11) =====
# Ключевой момент - создание ключа InprocServer32 с ПУСТЫМ значением по умолчанию

# Способ 1: через reg.exe (наиболее надёжный)
Try {
    # Создаём ключ CLSID
    reg add 'HKCU\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}' /f *> $null
    # Создаём InprocServer32 с ПУСТЫМ значением по умолчанию (/ve = default value, /d """" = empty string)
    reg add 'HKCU\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32' /f /ve /t REG_SZ /d """" *> $null
} Catch {}

# Способ 2: через PowerShell (дублируем для надежности)
Try {
    $clsidPath = 'HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}'
    $inprocPath = ""$clsidPath\InprocServer32""

    # Создаем оба ключа
    New-Item -Path $clsidPath -Force -ErrorAction SilentlyContinue | Out-Null
    New-Item -Path $inprocPath -Force -ErrorAction SilentlyContinue | Out-Null

    # Устанавливаем ПУСТОЕ значение по умолчанию
    New-ItemProperty -Path $inprocPath -Name '(Default)' -Value '' -PropertyType String -Force -ErrorAction SilentlyContinue | Out-Null
} Catch {}

# Включить отображение расширений файлов
Try { reg add 'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' /v HideFileExt /t REG_DWORD /d 0 /f | Out-Null } Catch {}

# Explorer будет перезапущен в конце всех операций

Write-Output 'Классическое меню включено.'

$global:LASTEXITCODE = 0
";

    private static string ScriptTelemetry() => @"
$ErrorActionPreference='SilentlyContinue'

# Службы телеметрии для отключения
$telemetryServices = @(
    'DiagTrack',                    # Connected User Experiences and Telemetry
    'dmwappushservice',             # WAP Push Message Routing Service
    'WMPNetworkSvc',                # Windows Media Player Network Sharing Service
    'PcaSvc',                       # Program Compatibility Assistant Service
    'diagsvc',                      # Diagnostic Execution Service
    'DPS',                          # Diagnostic Policy Service
    'WdiServiceHost',               # Diagnostic Service Host
    'WdiSystemHost',                # Diagnostic System Host
    'WerSvc',                       # Windows Error Reporting Service
    'wercplsupport',                # Problem Reports Control Panel Support
    'MapsBroker',                   # Downloaded Maps Manager
    'lfsvc',                        # Geolocation Service
    'RetailDemo',                   # Retail Demo Service
    'RemoteRegistry',               # Remote Registry
    'SysMain',                      # Superfetch (может быть полезен, но отключаем для POS)
    'TrkWks',                       # Distributed Link Tracking Client
    'WSearch',                      # Windows Search (если не нужен индекс)
    'XblAuthManager',               # Xbox Live Auth Manager
    'XblGameSave',                  # Xbox Live Game Save
    'XboxGipSvc',                   # Xbox Accessory Management Service
    'XboxNetApiSvc',                # Xbox Live Networking Service
    'DoSvc',                        # Delivery Optimization
    'WbioSrvc',                     # Windows Biometric Service
    'Fax',                          # Fax
    'SmsRouter',                    # SMS Router
    'PhoneSvc',                     # Phone Service
    'WalletService'                 # Wallet Service
)

foreach ($svc in $telemetryServices) {
    Try { Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue } Catch {}
    Try { Set-Service -Name $svc -StartupType Disabled -ErrorAction SilentlyContinue } Catch {}
    Try { sc.exe config $svc start= disabled *> $null } Catch {}
}

# Отключение телеметрии через реестр
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\DataCollection' /v AllowTelemetry /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\DataCollection' /v DisableTelemetry /t REG_DWORD /d 1 /f | Out-Null } Catch {}
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\DataCollection' /v MaxTelemetryAllowed /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\DataCollection' /v DoNotShowFeedbackNotifications /t REG_DWORD /d 1 /f | Out-Null } Catch {}
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\DataCollection' /v DisableEnterpriseAuthProxy /t REG_DWORD /d 1 /f | Out-Null } Catch {}

Try { reg add 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection' /v AllowTelemetry /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Policies\DataCollection' /v AllowTelemetry /t REG_DWORD /d 0 /f | Out-Null } Catch {}

# CEIP (Customer Experience Improvement Program)
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\SQMClient\Windows' /v CEIPEnable /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKLM\SOFTWARE\Microsoft\SQMClient\Windows' /v CEIPEnable /t REG_DWORD /d 0 /f | Out-Null } Catch {}

# Application Compatibility
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\AppCompat' /v AITEnable /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\AppCompat' /v DisableInventory /t REG_DWORD /d 1 /f | Out-Null } Catch {}
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\AppCompat' /v DisablePCA /t REG_DWORD /d 1 /f | Out-Null } Catch {}
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\AppCompat' /v DisableUAR /t REG_DWORD /d 1 /f | Out-Null } Catch {}

# Windows Error Reporting
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Error Reporting' /v Disabled /t REG_DWORD /d 1 /f | Out-Null } Catch {}
Try { reg add 'HKLM\SOFTWARE\Microsoft\Windows\Windows Error Reporting' /v Disabled /t REG_DWORD /d 1 /f | Out-Null } Catch {}
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\PCHealth\ErrorReporting' /v DoReport /t REG_DWORD /d 0 /f | Out-Null } Catch {}

# Advertising ID
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\AdvertisingInfo' /v DisabledByGroupPolicy /t REG_DWORD /d 1 /f | Out-Null } Catch {}
Try { reg add 'HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\AdvertisingInfo' /v Enabled /t REG_DWORD /d 0 /f | Out-Null } Catch {}

# Location tracking
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors' /v DisableLocation /t REG_DWORD /d 1 /f | Out-Null } Catch {}
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors' /v DisableLocationScripting /t REG_DWORD /d 1 /f | Out-Null } Catch {}
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors' /v DisableWindowsLocationProvider /t REG_DWORD /d 1 /f | Out-Null } Catch {}
Try { reg add 'HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location' /v Value /t REG_SZ /d 'Deny' /f | Out-Null } Catch {}

# Activity History
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\System' /v EnableActivityFeed /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\System' /v PublishUserActivities /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\System' /v UploadUserActivities /t REG_DWORD /d 0 /f | Out-Null } Catch {}

# Clipboard History
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\System' /v AllowClipboardHistory /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\System' /v AllowCrossDeviceClipboard /t REG_DWORD /d 0 /f | Out-Null } Catch {}

# Camera and Microphone privacy
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy' /v LetAppsAccessCamera /t REG_DWORD /d 2 /f | Out-Null } Catch {}
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy' /v LetAppsAccessMicrophone /t REG_DWORD /d 2 /f | Out-Null } Catch {}

# Delivery Optimization (P2P updates)
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\DeliveryOptimization' /v DODownloadMode /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config' /v DODownloadMode /t REG_DWORD /d 0 /f | Out-Null } Catch {}

# Handwriting recognition data
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\TabletPC' /v PreventHandwritingDataSharing /t REG_DWORD /d 1 /f | Out-Null } Catch {}
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\HandwritingErrorReports' /v PreventHandwritingErrorReports /t REG_DWORD /d 1 /f | Out-Null } Catch {}

# Inventory Collector
Try { reg add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\AppCompat' /v DisableInventory /t REG_DWORD /d 1 /f | Out-Null } Catch {}

# Speech, Inking, Typing
Try { reg add 'HKCU\SOFTWARE\Microsoft\Personalization\Settings' /v AcceptedPrivacyPolicy /t REG_DWORD /d 0 /f | Out-Null } Catch {}
Try { reg add 'HKCU\SOFTWARE\Microsoft\InputPersonalization' /v RestrictImplicitInkCollection /t REG_DWORD /d 1 /f | Out-Null } Catch {}
Try { reg add 'HKCU\SOFTWARE\Microsoft\InputPersonalization' /v RestrictImplicitTextCollection /t REG_DWORD /d 1 /f | Out-Null } Catch {}
Try { reg add 'HKCU\SOFTWARE\Microsoft\InputPersonalization\TrainedDataStore' /v HarvestContacts /t REG_DWORD /d 0 /f | Out-Null } Catch {}

# Tailored experiences
Try { reg add 'HKCU\SOFTWARE\Policies\Microsoft\Windows\CloudContent' /v DisableTailoredExperiencesWithDiagnosticData /t REG_DWORD /d 1 /f | Out-Null } Catch {}

# Diagnostics & feedback
Try { reg add 'HKCU\SOFTWARE\Microsoft\Siuf\Rules' /v NumberOfSIUFInPeriod /t REG_DWORD /d 0 /f | Out-Null } Catch {}

# Disable scheduled tasks related to telemetry
$telemetryTasks = @(
    '\Microsoft\Windows\Application Experience\Microsoft Compatibility Appraiser',
    '\Microsoft\Windows\Application Experience\ProgramDataUpdater',
    '\Microsoft\Windows\Application Experience\StartupAppTask',
    '\Microsoft\Windows\Autochk\Proxy',
    '\Microsoft\Windows\Customer Experience Improvement Program\Consolidator',
    '\Microsoft\Windows\Customer Experience Improvement Program\KernelCeipTask',
    '\Microsoft\Windows\Customer Experience Improvement Program\UsbCeip',
    '\Microsoft\Windows\DiskDiagnostic\Microsoft-Windows-DiskDiagnosticDataCollector',
    '\Microsoft\Windows\DiskDiagnostic\Microsoft-Windows-DiskDiagnosticResolver',
    '\Microsoft\Windows\Feedback\Siuf\DmClient',
    '\Microsoft\Windows\Feedback\Siuf\DmClientOnScenarioDownload',
    '\Microsoft\Windows\Maps\MapsToastTask',
    '\Microsoft\Windows\Maps\MapsUpdateTask',
    '\Microsoft\Windows\NetTrace\GatherNetworkInfo',
    '\Microsoft\Windows\PI\Sqm-Tasks',
    '\Microsoft\Windows\Power Efficiency Diagnostics\AnalyzeSystem',
    '\Microsoft\Windows\Shell\FamilySafetyMonitor',
    '\Microsoft\Windows\Shell\FamilySafetyRefreshTask',
    '\Microsoft\Windows\Windows Error Reporting\QueueReporting',
    '\Microsoft\Office\OfficeTelemetryAgentFallBack2016',
    '\Microsoft\Office\OfficeTelemetryAgentLogOn2016'
)

foreach ($task in $telemetryTasks) {
    Try { schtasks /Change /TN $task /Disable *> $null } Catch {}
    Try { Disable-ScheduledTask -TaskName $task -ErrorAction SilentlyContinue | Out-Null } Catch {}
}

$global:LASTEXITCODE = 0
";

    private static async Task WaitForDefenderWindowClosedAsync(IProgress<string>? progress, SetupFlowController controller)
    {
        var cancellationToken = controller.Token;
        progress?.Report("Открываю настройки Защитника Windows...");

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "windowsdefender://threatsettings",
                UseShellExecute = true
            });
        }
        catch
        {
            progress?.Report("⚠ Не удалось открыть настройки Defender");
            return;
        }

        await Task.Delay(1500, cancellationToken);

        var instructions = new[]
        {
            "",
            "!!! ВНИМАНИЕ: ТРЕБУЕТСЯ ДЕЙСТВИЕ !!!",
            "Отключите ВСЕ галочки в открывшемся окне:",
            "1. Защита в реальном времени (!) ВАЖНО",
            "2. Облачная защита",
            "3. Автоматическая отправка образцов",
            "4. Защита от подделки",
            "",
            "--> ЗАКРОЙТЕ ОКНО ЗАЩИТНИКА ДЛЯ ПРОДОЛЖЕНИЯ <--",
            ""
        };

        foreach (var line in instructions)
        {
            progress?.Report(line);
        }
        await Task.Delay(3000, cancellationToken);
        var processNames = new[] { "SecurityHealthUI", "WindowsSecurity", "SecHealthUI" };
        var windowTitles = new[] { "Безопасность Windows", "Windows Security", "Центр безопасности защитника Windows" };
        
        while (!HasVisibleWindow(processNames, windowTitles))
        {
            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report("⏳ Ожидание открытия окна Защитника...");
            await Task.Delay(2000, cancellationToken);
        }

        progress?.Report("✔ Окно Защитника открыто. Ожидание закрытия...");
        while (HasVisibleWindow(processNames, windowTitles))
        {
            await controller.WaitIfPausedAsync();
            cancellationToken.ThrowIfCancellationRequested();
            await Task.Delay(1000, cancellationToken);
        }
        
        progress?.Report("✔ Окно закрыто. Продолжаю...");
    }

    private static bool HasVisibleWindow(string[] processNames, string[] titles)
    {
        var pids = Process.GetProcesses()
            .Where(p => processNames.Contains(p.ProcessName, StringComparer.OrdinalIgnoreCase))
            .Select(p => p.Id)
            .ToHashSet();

        var found = false;
        EnumWindows((hwnd, lParam) =>
        {
            if (!IsWindowVisible(hwnd)) return true;
            GetWindowThreadProcessId(hwnd, out var pid);
            if (pids.Contains((int)pid))
            {
                found = true;
                return false;
            }
            var length = GetWindowTextLength(hwnd);
            if (length > 0)
            {
                var sb = new StringBuilder(length + 1);
                GetWindowText(hwnd, sb, sb.Capacity);
                var windowTitle = sb.ToString();
                if (titles.Any(t => string.Equals(windowTitle, t, StringComparison.OrdinalIgnoreCase)))
                {
                    found = true;
                    return false;
                }
            }

            return true;
        }, IntPtr.Zero);

        return found;
    }

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    private static string ScriptSslTls() => @"
$ErrorActionPreference='SilentlyContinue'

# ===== НАСТРОЙКА SSL/TLS ДЛЯ WINDOWS =====
# Решает проблему 'Could not create SSL/TLS secure channel'

# Создаём ключи реестра для TLS 1.1
New-Item -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1' -Force | Out-Null
New-Item -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Client' -Force | Out-Null
New-Item -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Server' -Force | Out-Null

# Включаем TLS 1.1
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Client' -Name 'DisabledByDefault' -Value 0 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Client' -Name 'Enabled' -Value 1 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Server' -Name 'DisabledByDefault' -Value 0 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.1\Server' -Name 'Enabled' -Value 1 -PropertyType DWord -Force | Out-Null

# Создаём ключи реестра для TLS 1.2
New-Item -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2' -Force | Out-Null
New-Item -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client' -Force | Out-Null
New-Item -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server' -Force | Out-Null

# Включаем TLS 1.2
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client' -Name 'DisabledByDefault' -Value 0 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client' -Name 'Enabled' -Value 1 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server' -Name 'DisabledByDefault' -Value 0 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server' -Name 'Enabled' -Value 1 -PropertyType DWord -Force | Out-Null

# Настройка .NET Framework для использования системных настроек TLS
# .NET 4.x
New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319' -Name 'SystemDefaultTlsVersions' -Value 1 -PropertyType DWord -Force -EA SilentlyContinue | Out-Null
New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319' -Name 'SchUseStrongCrypto' -Value 1 -PropertyType DWord -Force -EA SilentlyContinue | Out-Null
New-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319' -Name 'SystemDefaultTlsVersions' -Value 1 -PropertyType DWord -Force -EA SilentlyContinue | Out-Null
New-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319' -Name 'SchUseStrongCrypto' -Value 1 -PropertyType DWord -Force -EA SilentlyContinue | Out-Null

# .NET 2.x/3.x
New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\.NETFramework\v2.0.50727' -Name 'SystemDefaultTlsVersions' -Value 1 -PropertyType DWord -Force -EA SilentlyContinue | Out-Null
New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\.NETFramework\v2.0.50727' -Name 'SchUseStrongCrypto' -Value 1 -PropertyType DWord -Force -EA SilentlyContinue | Out-Null
New-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v2.0.50727' -Name 'SystemDefaultTlsVersions' -Value 1 -PropertyType DWord -Force -EA SilentlyContinue | Out-Null
New-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v2.0.50727' -Name 'SchUseStrongCrypto' -Value 1 -PropertyType DWord -Force -EA SilentlyContinue | Out-Null

Write-Output 'SSL/TLS настроен: TLS 1.1 и TLS 1.2 включены'

$global:LASTEXITCODE = 0
";
}
