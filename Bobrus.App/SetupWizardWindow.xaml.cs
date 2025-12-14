using System.Text.Json;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Runtime.InteropServices;
using System.Windows.Interop;
using System.Windows.Threading;
using System.Windows.Shell;
using System.Threading;
using Bobrus.App.Services;

namespace Bobrus.App;

public partial class SetupWizardWindow : Window
{
    private HwndSource? _hwndSource;
    private double _restoreLeft;
    private double _restoreTop;
    private double _restoreWidth;
    private double _restoreHeight;
    private readonly Thickness _defaultResizeBorder = new(8);
    private readonly WindowsSetupService _windowsSetupService = new();
    private LogService? _logger;
    private ReportService? _reportService;
    private DefenderTipWindow? _defenderTipWindow;
    private DateTime _setupStartUtc;
    private CancellationTokenSource _cts = new();
    private readonly List<PluginVersion> _selectedPlugins = new();

    public SetupWizardWindow()
    {
        InitializeComponent();
        _logger = new LogService(msg => Dispatcher.Invoke(() => AppendLog(msg, false)));
        _reportService = new ReportService();
        Loaded += OnLoaded;
        LocationChanged += OnWindowLocationChanged;
        SizeChanged += OnWindowSizeChanged;
    }

    private void AppendLog(string message, bool isError = false)
    {
        if (string.IsNullOrWhiteSpace(message)) return;
        var cleanMessage = message
            .Replace("✔ ", "")
            .Replace("✖ ", "")
            .Replace("⚠ ", "")
            .Replace("[PROGRESS]", "")
            .Trim();

        if (string.IsNullOrWhiteSpace(cleanMessage)) return;

        var prefix = isError ? "✖ " : "✔ ";
        if (message.Contains("✔") || message.Contains("✖") || message.Contains("⚠"))
        {
             prefix = "";
        }

        ResultBox.AppendText($"{prefix}{cleanMessage}{Environment.NewLine}");
        ResultBox.ScrollToEnd();
    }

    private void ShowDefenderTip()
    {
        Dispatcher.Invoke(() =>
        {
            if (_defenderTipWindow == null)
            {
                _defenderTipWindow = new DefenderTipWindow();
                _defenderTipWindow.Closed += (_, _) => _defenderTipWindow = null;
            }
            _defenderTipWindow.Show();
        });
    }

    private void HideDefenderTip()
    {
        Dispatcher.Invoke(() =>
        {
            _defenderTipWindow?.Close();
            _defenderTipWindow = null;
        });
    }

    private void OnAddPluginClicked(object sender, RoutedEventArgs e)
    {
        var window = new PluginSelectionWindow { Owner = this };
        if (window.ShowDialog() == true && window.AddedPlugins.Count > 0)
        {
            foreach (var plugin in window.AddedPlugins)
            {
                if (!_selectedPlugins.Any(p => p.Url == plugin.Url))
                {
                    _selectedPlugins.Add(plugin);
                }
            }
            UpdateSelectedPluginsUI();
        }
    }

    private void OnRemovePluginClicked(object sender, RoutedEventArgs e)
    {
        if (sender is System.Windows.Controls.Button btn && btn.Tag is PluginVersion plugin)
        {
            _selectedPlugins.Remove(plugin);
            UpdateSelectedPluginsUI();
        }
    }

    private void UpdateSelectedPluginsUI()
    {
        SelectedPluginsCountText.Text = $"Выбрано: {_selectedPlugins.Count}";
        SelectedPluginsList.ItemsSource = null;
        SelectedPluginsList.ItemsSource = _selectedPlugins;
    }

    private string _overlayDisplay = string.Empty;

    private void OnOverlayEnabledToggleChanged(object sender, RoutedEventArgs e)
    {
        if (OverlayFieldsPanel == null) return;
        OverlayFieldsPanel.Visibility = OverlayEnabledToggle.IsChecked == true 
            ? Visibility.Visible 
            : Visibility.Collapsed;
    }

    private void OnOverlayDisplayPickClicked(object sender, RoutedEventArgs e)
    {
        var picker = new ScreenPickerWindow(_overlayDisplay) { Owner = this };
        if (picker.ShowDialog() == true && !string.IsNullOrWhiteSpace(picker.SelectedDeviceName))
        {
            _overlayDisplay = picker.SelectedDeviceName!;
            OverlayDisplayLabel.Text = picker.SelectedDeviceName;
        }
    }

    public async void ResumeSetup()
    {
        var configPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Bobrus", "resume.json");
        if (!System.IO.File.Exists(configPath)) return;
        if (!MainWindow.IsAdministratorStatic())
        {
            AppendLog("Требуются права администратора. Перезапускаю...");
            MainWindow.RelaunchAsAdminStatic(launchSetupWizard: true);
            return;
        }
        await SetupAutorunForNewUserAsync();
        await Task.Delay(2000); 

        ShowCustomDialog("Ожидание загрузки", "Нажмите OK, когда система полностью загрузится, чтобы продолжить настройку.");

        try 
        {
            var json = await System.IO.File.ReadAllTextAsync(configPath);
            var options = JsonSerializer.Deserialize<WindowsSetupOptions>(json);
            if (options == null) return;
            ModeNew.IsChecked = true;
            OnModeChanged(this, new RoutedEventArgs());
            ModeExisting.IsEnabled = false;
            ModeNew.IsEnabled = false;
            ApplyButton.IsEnabled = false;
            CancelButton.IsEnabled = false;
            OptionsContainer.IsEnabled = false;
            _flowController = new SetupFlowController();
            _setupStartUtc = DateTime.UtcNow;
            ControlButtonsPanel.Visibility = Visibility.Visible;
            PauseButton.Content = "Пауза";

            await ExecuteSetup(options, _flowController); 
        }
        catch (Exception ex)
        {
             System.Windows.MessageBox.Show($"Ошибка возобновления: {ex.Message}");
        }
    }
    private async Task SetupAutorunForNewUserAsync()
    {
        try
        {
            AppendLog("Phase 2: Настройка автозапуска для нового пользователя...");
            
            var userName = Environment.UserName;
            var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var bobrusDesktopPath = System.IO.Path.Combine(desktopPath, "Bobrus.exe");
            var rhelperPath = @"C:\Program Files (x86)\Rhelper39\RHusr_v39.exe";
            if (!System.IO.File.Exists(bobrusDesktopPath))
            {
                var usersDir = @"C:\Users";
                string? bobrusSource = null;
                
                foreach (var userDir in System.IO.Directory.GetDirectories(usersDir))
                {
                    if (userDir.Contains(userName)) continue; 
                    
                    var searchPaths = new[]
                    {
                        System.IO.Path.Combine(userDir, "Desktop", "Bobrus.exe"),
                        System.IO.Path.Combine(userDir, "Downloads", "Bobrus.exe"),
                        System.IO.Path.Combine(userDir, "Рабочий стол", "Bobrus.exe")
                    };
                    
                    foreach (var path in searchPaths)
                    {
                        if (System.IO.File.Exists(path))
                        {
                            bobrusSource = path;
                            break;
                        }
                    }
                    if (bobrusSource != null) break;
                }
                
                if (bobrusSource != null)
                {
                    System.IO.File.Copy(bobrusSource, bobrusDesktopPath, true);
                    AppendLog($"Bobrus.exe скопирован на рабочий стол: {bobrusDesktopPath}");
                }
                else
                {
                    AppendLog("ВНИМАНИЕ: Bobrus.exe не найден для копирования");
                }
            }
            if (System.IO.File.Exists(bobrusDesktopPath))
            {
                using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true);
                key?.SetValue("Bobrus", $"\"{bobrusDesktopPath}\" --start-hidden");
                AppendLog($"Bobrus добавлен в автозапуск (свёрнутый): {bobrusDesktopPath}");
            }
            if (System.IO.File.Exists(rhelperPath))
            {
                using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true);
                key?.SetValue("Rhelper", $"\"{rhelperPath}\"");
                AppendLog($"Rhelper добавлен в автозапуск: {rhelperPath}");
            }
            else
            {
                AppendLog($"Rhelper не найден: {rhelperPath}");
            }
            
            await Task.CompletedTask;
        }
        catch (Exception ex)
        {
            AppendLog($"Ошибка настройки автозапуска: {ex.Message}");
        }
    }

    protected override void OnClosed(EventArgs e)
    {
        base.OnClosed(e);
        HideDefenderTip();
        _cts.Cancel();
        _cts.Dispose();
    }

    private void OnLoaded(object? sender, RoutedEventArgs e)
    {
        var workArea = SystemParameters.WorkArea;
        MaxHeight = workArea.Height;
        MaxWidth = workArea.Width;
        CaptureRestoreBounds();
    }

    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);
        AdjustToWorkArea();
        AdjustResizeBorder();
        _hwndSource = PresentationSource.FromVisual(this) as HwndSource;
        _hwndSource?.AddHook(WndProc);
    }

    protected override void OnStateChanged(EventArgs e)
    {
        base.OnStateChanged(e);
        AdjustToWorkArea();
        AdjustResizeBorder();
    }

    private void OnTopBarMouseDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton != MouseButton.Left)
        {
            return;
        }

        if (e.ClickCount == 2)
        {
            ToggleMaximize();
            return;
        }

        if (WindowState == WindowState.Maximized)
        {
            StartDragFromMaximized(e);
            return;
        }

        DragMove();
    }

    private void OnMinimizeClicked(object sender, RoutedEventArgs e)
    {
        WindowState = WindowState.Minimized;
    }

    private void OnToggleMaximizeClicked(object sender, RoutedEventArgs e)
    {
        ToggleMaximize();
    }

    private void OnCloseClicked(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void OnModeChanged(object sender, RoutedEventArgs e)
    {
        if (NewUserGroup == null) return;
        NewUserGroup.Visibility = (ModeNew.IsChecked == true) ? Visibility.Visible : Visibility.Collapsed;
    }

    private void OnTerminalTextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        UpdateComputerNamePreview();
    }

    private void OnTerminalTypeRadioChanged(object sender, RoutedEventArgs e)
    {
        UpdateIdHint();
        UpdateComputerNamePreview();
    }

    private void UpdateIdHint()
    {
        if (IdHintLabel == null) return;
        var (prefix, _, _) = GetTerminalTypeInfo();
        IdHintLabel.Text = $"ID ({prefix}XX):";
    }

    private (string prefix, string shortName, string fullName) GetTerminalTypeInfo()
    {
        if (TypeGK?.IsChecked == true) return ("10", "ГК", "Главная касса");
        if (TypeVedomy?.IsChecked == true) return ("20", "Ведомый терминал", "Ведомый терминал");
        if (TypeKitchen?.IsChecked == true) return ("30", "Кухонный экран", "Кухонный экран");
        if (TypeKiosk?.IsChecked == true) return ("40", "Киоск", "Киоск");
        if (TypeOffice?.IsChecked == true) return ("50", "Офисный ПК", "Офисный ПК");
        return ("10", "ГК", "Главная касса");
    }

    private void UpdateComputerNamePreview()
    {
        if (ComputerNamePreview == null || TerminalIdBox == null) return;
        ComputerNamePreview.Text = GenerateSystemComputerName();
    }
    private string GenerateSystemComputerName()
    {
        var id = TerminalIdBox?.Text?.PadLeft(2, '0') ?? "01";
        var (prefix, _, _) = GetTerminalTypeInfo();
        string typeCode;
        if (TypeGK?.IsChecked == true) typeCode = "GK";
        else if (TypeVedomy?.IsChecked == true) typeCode = "VED";
        else if (TypeKitchen?.IsChecked == true) typeCode = "KITCHEN";
        else if (TypeKiosk?.IsChecked == true) typeCode = "KIOSK";
        else if (TypeOffice?.IsChecked == true) typeCode = "OFFICE";
        else typeCode = "GK";
        
        return $"{typeCode}-{prefix}{id}";
    }

    private void OnSelectAllWindowsOptions(object sender, RoutedEventArgs e)
    {
        SetAllWindowsOptions(true);
    }

    private void OnDeselectAllWindowsOptions(object sender, RoutedEventArgs e)
    {
        SetAllWindowsOptions(false);
    }

    private void SetAllWindowsOptions(bool isChecked)
    {
        foreach (var toggle in GetOptionToggles())
        {
            if (toggle.IsEnabled)
            {
                toggle.IsChecked = isChecked;
            }
        }
    }

    private void OnSelectAllSoftware(object sender, RoutedEventArgs e)
    {
        SetAllSoftwareOptions(true);
    }

    private void OnDeselectAllSoftware(object sender, RoutedEventArgs e)
    {
        SetAllSoftwareOptions(false);
    }

    private void SetAllSoftwareOptions(bool isChecked)
    {
        foreach (var toggle in GetSoftwareToggles())
        {
            if (toggle.IsEnabled)
            {
                toggle.IsChecked = isChecked;
            }
        }
    }

    private IEnumerable<System.Windows.Controls.CheckBox> GetSoftwareToggles()
    {
        yield return Sw7ZipToggle;
        yield return SwNotepadToggle;
        yield return SwBraveToggle;
        yield return SwVlcToggle;
        yield return SwDotNet48Toggle;
        yield return SwRhelperToggle;
        yield return SwRustDeskToggle;
        yield return SwAnyDeskToggle;
        yield return SwAssistantToggle;
        yield return SwAtolDriverToggle;
        yield return SwShtrihDriverToggle;
        yield return SwOrderCheckToggle;
        yield return SwClearBatToggle;
        yield return SwFrontToolsToggle;
        yield return SwFrontToolsSqliteToggle;
        yield return SwZabbixToggle;
        yield return SwComPortCheckerToggle;
        yield return SwDatabaseNetToggle;
        yield return SwAdvancedIpScannerToggle;
        yield return SwPrinterTestToggle;
    }

    private void OnCancelClicked(object sender, RoutedEventArgs e)
    {
        var oldCts = _cts;
        oldCts.Cancel();
        _cts = new CancellationTokenSource();
        try { oldCts.Dispose(); } catch { } 
        
        AppendLog("⚠ Операция отменена пользователем");
        StepProgress.IsIndeterminate = false;
        StepProgress.Value = 0;
        StepStatusText.Text = "Отменено";
        
        RestoreUIAfterSetup();
    }

    private void RestoreUIAfterSetup()
    {
        ControlButtonsPanel.Visibility = Visibility.Collapsed;
        _flowController?.Dispose();
        _flowController = null;
        
        ModeExisting.IsEnabled = true;
        ModeNew.IsEnabled = true;
        NewUserGroup.IsEnabled = true;
        ApplyButton.Visibility = Visibility.Visible;
        ApplyButton.IsEnabled = true;
        CancelButton.Visibility = Visibility.Collapsed;
        OptionsContainer.IsEnabled = true;
    }
    private SetupFlowController? _flowController;

    private void OnPauseClicked(object sender, RoutedEventArgs e)
    {
        if (_flowController == null) return;
        
        if (_flowController.IsPaused)
        {
            _flowController.Resume();
            _reportService?.LogPause(false);
            PauseButton.Content = "Пауза";
            AppendLog("▶ Продолжаем...");
        }
        else
        {
            _flowController.Pause();
            _reportService?.LogPause(true);
            PauseButton.Content = "Продолжить";
            AppendLog("⏸ Пауза...");
        }
    }

    private void OnStopClicked(object sender, RoutedEventArgs e)
    {
        if (ShowConfirmationDialog("Отмена", "Прервать процесс пусконаладки?"))
        {
             
             _flowController?.Cancel();
             _reportService?.LogStop();
             AppendLog("🛑 Остановка процесса...");
        }
    }

    private async void OnApplyClicked(object sender, RoutedEventArgs e)
    {
        if (ModeNew.IsChecked == true)
        {
            if (!int.TryParse(TerminalIdBox.Text, out var id) || id < 1 || id > 99)
            {
                ShowCustomDialog("Ошибка", "Введите ID терминала (01-99)!", isError: true);
                TerminalIdBox.Focus();
                return;
            }
        }
        if ((IikoFrontToggle.IsChecked == true || IikoOfficeToggle.IsChecked == true || IikoChainToggle.IsChecked == true)
            && string.IsNullOrWhiteSpace(IikoServerBox.Text))
        {
            ShowCustomDialog("Ошибка", "Введите адрес сервера IIKO!", isError: true);
            IikoServerBox.Focus();
            return;
        }
        if (!ShowConfirmationDialog("Подтверждение", "Начать процесс пусконаладки?\n\nВсе выбранные настройки будут применены к системе."))
        {
            return;
        }

        if (!MainWindow.IsAdministratorStatic())
        {
            AppendLog("Требуются права администратора. Перезапускаю...");
            MainWindow.RelaunchAsAdminStatic(launchSetupWizard: true);
            return;
        }

        ResultBox.Clear();
        _setupStartUtc = DateTime.UtcNow;
        ModeExisting.IsEnabled = false;
        ModeNew.IsEnabled = false;
        NewUserGroup.IsEnabled = false;
        ApplyButton.Visibility = Visibility.Collapsed;
        CancelButton.Visibility = Visibility.Collapsed; 
        OptionsContainer.IsEnabled = false;
        ControlButtonsPanel.Visibility = Visibility.Visible;
        PauseButton.Content = "Пауза";

        _flowController = new SetupFlowController();
        var ct = _flowController.Token;
        var rhelperService = new RhelperSetupService();
            var rhelperProgress = new Progress<string>(line => 
            {
                Dispatcher.Invoke(() => 
                {
                    if (line.StartsWith("[PROGRESS]"))
                    {
                        if (int.TryParse(line.Replace("[PROGRESS]", ""), out int p))
                        {
                            StepProgress.Value = p;
                        }
                    }
                    else
                    {
                        AppendLog(line);
                    }
                });
            });
            
            StepProgress.Value = 0;
            StepStatusText.Text = "Подготовка удалённого доступа...";
            bool rhelperDefenderDisabled = false;
            bool shouldInstallRhelper = (ModeNew.IsChecked == true && RhelperToggle.IsChecked == true);
            
            if (shouldInstallRhelper)
            {
                try
                {
                    await _flowController.WaitIfPausedAsync();
                    ShowDefenderTip();
                    await rhelperService.WaitForDefenderDisabledAsync(rhelperProgress, _flowController); 
                    HideDefenderTip();
                    rhelperDefenderDisabled = true;
                    
                    await _flowController.WaitIfPausedAsync();
                    StepProgress.Visibility = Visibility.Visible;
                    StepProgress.IsIndeterminate = false;
                    StepProgress.Value = 0;
                    
                    var installerPath = await rhelperService.DownloadAsync(rhelperProgress, ct);
                    
                    if (installerPath == null)
                    {
                         ShowCustomDialog("Ошибка", "Не удалось скачать Rhelper", isError: true);
                         RestoreUIAfterSetup();
                         return;
                    }
                    
                    await _flowController.WaitIfPausedAsync();
                    StepProgress.IsIndeterminate = true; 
                    var installProcess = rhelperService.StartInstaller(installerPath, rhelperProgress);
                    
                    if (installProcess == null)
                    {
                        ShowCustomDialog("Ошибка", "Не удалось запустить установщик Rhelper", isError: true);
                        RestoreUIAfterSetup();
                        return;
                    }
                    
                    try { if (File.Exists(installerPath)) File.Delete(installerPath); } catch { }
                    var showIdTask = rhelperService.WaitForShowIdClosedAsync(rhelperProgress, _flowController); 
                    var installExitTask = installProcess.WaitForExitAsync(ct);
                    
                    var completedTask = await Task.WhenAny(showIdTask, installExitTask);
                    
                    if (completedTask == installExitTask && !showIdTask.IsCompleted)
                    {
                         try
                        {
                            var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
                            var showIdExe = Path.Combine(programFiles, "Rhelper39", "show-id.exe");
                            var mainExe = Path.Combine(programFiles, "Rhelper39", "RHurs_v39.exe");
                            
                            if (File.Exists(showIdExe)) Process.Start(new ProcessStartInfo { FileName = showIdExe, UseShellExecute = true });
                            else if (File.Exists(mainExe)) Process.Start(new ProcessStartInfo { FileName = mainExe, UseShellExecute = true });
                        }
                        catch { }
                        
                        await showIdTask;
                    }
                    else
                    {
                        await showIdTask;
                    }

                    StepProgress.Visibility = Visibility.Collapsed;
                }
                catch (OperationCanceledException)
                {
                    AppendLog("Отменено пользователем.");
                    RestoreUIAfterSetup();
                    return;
                }
                catch (Exception ex)
                {
                    ShowCustomDialog("Ошибка", $"Ошибка Rhelper: {ex.Message}", isError: true);
                    RestoreUIAfterSetup();
                    return;
                }
            }
        
        await _flowController.WaitIfPausedAsync();

        AppendLog("");
        AppendLog("===== Начало пусконаладки =====");
        AppendLog("");

        var options = new WindowsSetupOptions
        {
            StartMenu = StartMenuToggle.IsChecked == true,
            UsbPower = UsbPowerToggle.IsChecked == true,
            Sleep = SleepToggle.IsChecked == true,
            PowerPlan = PowerPlanToggle.IsChecked == true,
            Defender = DefenderToggle.IsChecked == true,
            SslTls = SslTlsToggle.IsChecked == true,
            Bloat = BloatToggle.IsChecked == true,
            Theme = ThemeToggle.IsChecked == true,
            Uac = UacToggle.IsChecked == true,
            Explorer = ExplorerToggle.IsChecked == true,
            Accessibility = AccessibilityToggle.IsChecked == true,
            Lock = LockToggle.IsChecked == true,
            Locale = LocaleToggle.IsChecked == true,
            ContextMenu = ContextMenuToggle.IsChecked == true,
            Telemetry = TelemetryToggle.IsChecked == true,
            Cleanup = CleanupToggle.IsChecked == true,
            CreateNewUser = ModeNew.IsChecked == true,
            ComputerName = GenerateSystemComputerName(),
            DeleteOldUser = DeleteOldUserToggle.IsChecked == true,
            IikoServerUrl = IikoServerBox.Text.Trim(),
            IikoFront = IikoFrontToggle.IsChecked == true,
            IikoOffice = IikoOfficeToggle.IsChecked == true,
            IikoChain = IikoChainToggle.IsChecked == true,
            IikoCard = IikoCardToggle.IsChecked == true,
            IikoFrontAutostart = IikoFrontAutostartToggle.IsChecked == true,
            IikoHandCardRoll = IikoHandCardRollToggle.IsChecked == true,
            IikoMinimizeButton = IikoMinimizeButtonToggle.IsChecked == true,
            IikoSetServerUrl = IikoConfigServerToggle.IsChecked == true,
            IikoPlugins = _selectedPlugins,
            SkipDefenderPrompt = rhelperDefenderDisabled
        };
        try
        {
            var configPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Bobrus", "resume.json");
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(configPath)!);
            var json = JsonSerializer.Serialize(options, new JsonSerializerOptions { WriteIndented = true });
            await File.WriteAllTextAsync(configPath, json);
        }
        catch (Exception ex)
        {
            _logger?.Error($"Failed to save resume config: {ex.Message}");
        }
        if (_reportService == null) _reportService = new Services.ReportService();
        
        _reportService.Start();
        _reportService.LogOptions(options);

        await ExecuteSetup(options, _flowController);
    }

    private async Task ExecuteSetup(WindowsSetupOptions options, SetupFlowController controller)
    {
        ApplyButton.IsEnabled = false;
        OptionsContainer.IsEnabled = false;

        try
        {
            var enabledCount = options.GetType().GetProperties()
                .Where(p => p.PropertyType == typeof(bool) && p.Name != "CreateNewUser") 
                .Count(p => p.GetValue(options) is true);
            
            if (options.CreateNewUser) enabledCount++;

            StepProgress.Value = 0;
            StepStatusText.Text = "Подготовка...";
            
            _logger?.Info("=== Начало настройки Windows ===");
            
            var completedSteps = 0;
            var progress = new Progress<string>(line =>
            {
                if (line.StartsWith("[VERBOSE] "))
                {
                     _logger?.Verbose(line.Substring(10));
                     return;
                }

                _logger?.Info(line);
                if (line.Contains("Открываю настройки Защитника", StringComparison.OrdinalIgnoreCase))
                {
                    ShowDefenderTip();
                }
                if (line.Contains("Окно закрыто", StringComparison.OrdinalIgnoreCase))
                {
                    HideDefenderTip();
                }
                if (TryParseStepStart(line, enabledCount, out var currentStep, out var totalSteps, out var title))
                {
                    _reportService?.AddStep(title, "Running");
                }
                if (IsStepCompleted(line))
                {
                    var isSuccess = !line.StartsWith("✖");
                    _reportService?.UpdateStep("Current", isSuccess ? "Success" : "Error", line);
                }
                var iikoMatch = System.Text.RegularExpressions.Regex.Match(line, @"\[IIKO\] Скачивание .+?: (\d+)%");
                if (iikoMatch.Success)
                {
                    var percent = int.Parse(iikoMatch.Groups[1].Value);
                    StepProgress.Value = percent;
                    StepStatusText.Text = $"Скачивание iiko: {percent}%";
                    return;
                }
                if (line.Contains("[IIKO] Установка"))
                {
                    StepProgress.IsIndeterminate = true;
                    StepStatusText.Text = line.Replace("[IIKO] ", "");
                    return;
                }
                if (line.Contains("✔ [IIKO]") || line.Contains("⚠ [IIKO]"))
                {
                    StepProgress.IsIndeterminate = false;
                    StepProgress.Value = 100;
                    return;
                }

                if (TryParseStepStart(line, enabledCount, out var currentStep2, out var totalSteps2, out var title2))
                {
                    StepProgress.IsIndeterminate = false;
                    completedSteps = Math.Min(completedSteps, currentStep2 - 1);
                    var percent = totalSteps2 > 0 ? Math.Max(0, ((currentStep2 - 1) * 100.0 / totalSteps2)) : 0;
                    StepProgress.Value = percent;
                    StepStatusText.Text = $"Выполняется {currentStep2}/{totalSteps2}: {title2}";
                    return;
                }

                if (IsStepCompleted(line))
                {
                    completedSteps = Math.Min(enabledCount, completedSteps + 1);
                    var percent = enabledCount > 0 ? completedSteps * 100.0 / enabledCount : 0;
                    StepProgress.Value = percent;
                    StepStatusText.Text = $"Выполнено {completedSteps}/{enabledCount} ({percent:F0}%)";
                    return;
                }
            });
            Func<Task>? softwareCallback = null;
            var isPhase1 = options.CreateNewUser && 
                !string.Equals(Environment.UserName, "POS-терминал", StringComparison.OrdinalIgnoreCase);
            if (!isPhase1)
            {
                softwareCallback = async () => await InstallSelectedSoftwareAsync(_cts.Token);
            }
            
            var result = await _windowsSetupService.ApplyAsync(options, progress, controller, softwareCallback);
            StepProgress.IsIndeterminate = false;
            StepProgress.Value = 100;
            
            if (result.TotalSteps > 0)
            {
                var donePercent = result.DoneSteps * 100.0 / result.TotalSteps;
                StepStatusText.Text = $"Выполнено {result.DoneSteps}/{result.TotalSteps} ({donePercent:F0}%)";
            }
            else
            {
                StepStatusText.Text = "✔ Установка завершена";
            }
            
            _logger?.Info("=== Настройка завершена ===");

            if (!result.Success && result.DoneSteps == 0)
            {
                _logger?.Error("Пусконаладка прервана из-за ошибок.");
                RestoreUIAfterSetup(); 
            }
            else
            {
                if (!result.Success)
                {
                   _logger?.Error("Часть шагов завершилась с ошибкой.");
                }
                if (result.Output.Contains("Выход из системы"))
                {
                }
                else
                {
                    CompletionOverlay.Visibility = Visibility.Visible;
                }
            }
            if (!result.Output.Contains("Выход из системы"))
            {
                await SaveOverlaySettingsAsync();
            }
        }
        catch (OperationCanceledException)
        {
            _logger?.Info("✖ Операция отменена.");
            CancellationOverlay.Visibility = Visibility.Visible;
        }
        catch (Exception ex)
        {
            _logger?.Error($"✖ Ошибка: {ex.Message}");
        }
        finally
        {
            try 
            { 
                 _reportService?.Finish(); 
                 _reportService?.GenerateReport(options?.ComputerName ?? Environment.MachineName); 
            } 
            catch (Exception ex) 
            { 
                _logger?.Error($"Failed to generate report: {ex.Message}"); 
            }

            HideDefenderTip(); 
            if (CompletionOverlay.Visibility != Visibility.Visible && CancellationOverlay.Visibility != Visibility.Visible)
            {
                RestoreUIAfterSetup();
            }
        }
    }
    

    private void OnRebootYesClicked(object sender, RoutedEventArgs e)
    {
        try
        {
            Process.Start("shutdown", "/r /t 0");
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"Не удалось перезагрузить ПК: {ex.Message}");
        }
    }

    private void OnRebootNoClicked(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void ToggleMaximize()
    {
        if (WindowState == WindowState.Maximized)
        {
            WindowState = WindowState.Normal;
            return;
        }

        var workArea = SystemParameters.WorkArea;
        MaxHeight = workArea.Height;
        MaxWidth = workArea.Width;
        WindowState = WindowState.Maximized;
    }

    private void AdjustToWorkArea()
    {
        var workArea = GetMonitorWorkArea();

        MaxHeight = workArea.Height;
        MaxWidth = workArea.Width;

        if (WindowState == WindowState.Maximized)
        {
            Left = workArea.Left;
            Top = workArea.Top;
            Width = workArea.Width;
            Height = workArea.Height;
            return;
        }

        Width = Math.Min(Width, workArea.Width);
        Height = Math.Min(Height, workArea.Height);

        if (Left < workArea.Left) Left = workArea.Left;
        if (Top < workArea.Top) Top = workArea.Top;
        if (Left + Width > workArea.Right) Left = workArea.Right - Width;
        if (Top + Height > workArea.Bottom) Top = workArea.Bottom - Height;
    }

    private Rect GetMonitorWorkArea()
    {
        var handle = new WindowInteropHelper(this).Handle;
        var monitor = MonitorFromWindow(handle, MonitorDefaultToNearest);

        var info = new MONITORINFO
        {
            cbSize = Marshal.SizeOf<MONITORINFO>()
        };

        if (!GetMonitorInfo(monitor, ref info))
        {
            var area = SystemParameters.WorkArea;
            return new Rect(area.Left, area.Top, area.Width, area.Height);
        }

        var (scaleX, scaleY) = GetDpiScale();
        var work = info.rcWork;
        var left = work.Left / scaleX;
        var top = work.Top / scaleY;
        var width = (work.Right - work.Left) / scaleX;
        var height = (work.Bottom - work.Top) / scaleY;
        return new Rect(left, top, width, height);
    }

    private (double scaleX, double scaleY) GetDpiScale()
    {
        var source = PresentationSource.FromVisual(this);
        if (source?.CompositionTarget is { } target)
        {
            return (target.TransformToDevice.M11, target.TransformToDevice.M22);
        }

        return (1.0, 1.0);
    }

    private IEnumerable<System.Windows.Controls.CheckBox> GetOptionToggles()
    {
        yield return StartMenuToggle;
        yield return UsbPowerToggle;
        yield return SleepToggle;
        yield return PowerPlanToggle;
        yield return DefenderToggle;
        yield return SslTlsToggle;
        yield return BloatToggle;
        yield return ThemeToggle;
        yield return UacToggle;
        yield return ExplorerToggle;
        yield return AccessibilityToggle;
        yield return LockToggle;
        yield return LocaleToggle;
        yield return ContextMenuToggle;
        yield return TelemetryToggle;
        yield return CleanupToggle;
    }

    private async Task InstallSelectedSoftwareAsync(CancellationToken ct)
    {
        var softwareToInstall = GetSelectedSoftwareItems().ToList();
        if (softwareToInstall.Count == 0) return;

        AppendLog("");
        AppendLog("===== Установка программ =====");
        AppendLog("");

        var service = new Services.SoftwareInstallService();
        var total = softwareToInstall.Count;
        Task<string?>? nextDownloadTask = null;
        int nextDownloadIndex = -1;

        for (int i = 0; i < total; i++)
        {
            if (ct.IsCancellationRequested) break;
            await _flowController.WaitIfPausedAsync();

            var item = softwareToInstall[i];
            var current = i + 1;
            
            StepStatusText.Text = $"Скачивание: {item.Name} ({current}/{total})";
            StepProgress.Value = (i * 100.0) / total;
            AppendLog($"[{current}/{total}] {item.Name}");
            
            string? filePath = null;
            if (nextDownloadTask != null && nextDownloadIndex == i)
            {
                AppendLog($"   ⏳ Ожидание завершения скачивания...");
                filePath = await nextDownloadTask;
                nextDownloadTask = null;
            }
            if (filePath == null)
            {
                try
                {
                    AppendLog($"   ⬇ Скачивание...");
                    filePath = await service.DownloadOnlyAsync(item, ct);
                }
                catch (Exception ex)
                {
                    AppendLog($"   ✖ Ошибка скачивания: {ex.Message}");
                    continue;
                }
            }
            if (i + 1 < total)
            {
                var nextItem = softwareToInstall[i + 1];
                nextDownloadIndex = i + 1;
                nextDownloadTask = Task.Run(async () =>
                {
                    try
                    {
                        return await service.DownloadOnlyAsync(nextItem, ct);
                    }
                    catch
                    {
                        return null;
                    }
                }, ct);
            }
            StepStatusText.Text = $"Установка: {item.Name} ({current}/{total})";
            AppendLog($"   📦 Установка...");
            
            try
            {
                await service.InstallFromFileAsync(item, filePath, ct);
                AppendLog($"   ✔ Готово");
            }
            catch (OperationCanceledException)
            {
                AppendLog($"   ⚠ Отменено");
                break;
            }
            catch (Exception ex)
            {
                AppendLog($"   ✖ Ошибка установки: {ex.Message}");
            }

            StepProgress.Value = ((i + 1) * 100.0) / total;
        }

        StepProgress.Value = 100;
        StepStatusText.Text = $"Установлено {total} программ";
    }

    private IEnumerable<Services.SoftwareItem> GetSelectedSoftwareItems()
    {
        if (Sw7ZipToggle.IsChecked == true) yield return new Services.SoftwareItem("7-Zip", "https://github.com/Feuda1/Programs-for-Bobrik/releases/download/v1.0.0/7z2500-x64.exe", Services.SoftwareInstallType.SilentExe, "/S");
        if (SwNotepadToggle.IsChecked == true) yield return new Services.SoftwareItem("Notepad++", "https://github.com/Feuda1/Programs-for-Bobrik/releases/download/v1.0.0/npp.8.8.2.Installer.x64.exe", Services.SoftwareInstallType.SilentExe, "/S");
        if (SwBraveToggle.IsChecked == true) yield return new Services.SoftwareItem("Google Chrome", "https://dl.google.com/chrome/install/latest/chrome_installer.exe", Services.SoftwareInstallType.SilentExe, "/silent /install");
        if (SwVlcToggle.IsChecked == true) yield return new Services.SoftwareItem("VLC", "https://github.com/Feuda1/Programs-for-Bobrik/releases/download/v1.0.0/vlc-3.0.21-win64.exe", Services.SoftwareInstallType.SilentExe, "/S");
        if (SwDotNet48Toggle.IsChecked == true) yield return new Services.SoftwareItem(".NET 4.8", "https://github.com/Feuda1/Programs-for-Bobrik/releases/download/v1.0.0/ndp48-web.exe", Services.SoftwareInstallType.RunVisible);
        if (SwRhelperToggle.IsChecked == true) yield return new Services.SoftwareItem("Rhelper", "https://repo.denvic.ru/remote-access/remote-access-setup.exe", Services.SoftwareInstallType.SilentExe, "/PASSWORD=\"remote-access-setup\" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-");
        if (SwRustDeskToggle.IsChecked == true) yield return new Services.SoftwareItem("RustDesk", "https://github.com/rustdesk/rustdesk/releases/download/1.4.4/rustdesk-1.4.4-x86_64.exe", Services.SoftwareInstallType.SilentExe, "--silent-install");
        if (SwAnyDeskToggle.IsChecked == true) yield return new Services.SoftwareItem("AnyDesk", "https://download.anydesk.com/AnyDesk.exe", Services.SoftwareInstallType.RevealOnly);
        if (SwAssistantToggle.IsChecked == true) yield return new Services.SoftwareItem("Ассистент", "https://github.com/Feuda1/Programs-for-Bobrik/releases/download/v1.0.0/assistant_install_6.exe", Services.SoftwareInstallType.RunVisible);
        if (SwAtolDriverToggle.IsChecked == true) yield return new Services.SoftwareItem("Драйвер АТОЛ 10.10.7", "https://github.com/Feuda1/Programs-for-Bobrik/releases/download/v1.0.0/KKT10-10.10.7.0-windows32-setup.exe", Services.SoftwareInstallType.RunVisible);
        if (SwShtrihDriverToggle.IsChecked == true) yield return new Services.SoftwareItem("Драйвер Штрих 5.21.1207", "https://github.com/Feuda1/Programs-for-Bobrik/releases/download/v1.0.0/Poscenter_DrvKKT_5.21_1207_x64.exe", Services.SoftwareInstallType.SilentExe, "/S");
        if (SwOrderCheckToggle.IsChecked == true) yield return new Services.SoftwareItem("OrderCheck", "https://clearbat.iiko.online/downloads/OrderCheck.exe", Services.SoftwareInstallType.RevealOnly);
        if (SwClearBatToggle.IsChecked == true) yield return new Services.SoftwareItem("CLEAR.bat", "https://clearbat.iiko.online/downloads/CLEAR.bat.exe", Services.SoftwareInstallType.SilentExe, "/S");
        if (SwFrontToolsToggle.IsChecked == true) yield return new Services.SoftwareItem("FrontTools", "https://fronttools.iiko.it/FrontTools.exe", Services.SoftwareInstallType.RevealOnly);
        if (SwFrontToolsSqliteToggle.IsChecked == true) yield return new Services.SoftwareItem("FrontTools SQLite", "https://fronttools.iiko.it/fronttools_sqlite.zip", Services.SoftwareInstallType.ExtractZip);
        if (SwZabbixToggle.IsChecked == true) yield return new Services.SoftwareItem("Zabbix Agent", "https://github.com/Feuda1/Programs-for-Bobrik/releases/download/v1.0.0/INSTALLZabbix.exe", Services.SoftwareInstallType.SilentExe, "/S");
        if (SwComPortCheckerToggle.IsChecked == true) yield return new Services.SoftwareItem("ComPort Checker", "https://github.com/Feuda1/Programs-for-Bobrik/releases/download/v1.0.0/ComPortChecker.1.1.zip", Services.SoftwareInstallType.ExtractZip);
        if (SwDatabaseNetToggle.IsChecked == true) yield return new Services.SoftwareItem("Database.NET", "https://fishcodelib.com/files/DatabaseNet4.zip", Services.SoftwareInstallType.ExtractZip);
        if (SwAdvancedIpScannerToggle.IsChecked == true) yield return new Services.SoftwareItem("Advanced IP Scanner", "https://download.advanced-ip-scanner.com/download/files/Advanced_IP_Scanner_2.5.4594.1.exe", Services.SoftwareInstallType.SilentExe, "/S");
        if (SwPrinterTestToggle.IsChecked == true) yield return new Services.SoftwareItem("Printer Test", "https://725920.selcdn.ru/upload_portkkm/iblock/929/9291b0d0c76c7c5756b81732df929086/Printer-TEST-V3.1C.zip", Services.SoftwareInstallType.ExtractZip);
    }

    private async Task SaveOverlaySettingsAsync()
    {
        if (OverlayEnabledToggle.IsChecked != true) 
        {
            _logger?.Info("Оверлей CRM/Касса выключен, пропускаем сохранение.");
            return;
        }
        
        var crm = OverlayCrmBox.Text?.Trim() ?? string.Empty;
        var cash = OverlayCashBox.Text?.Trim() ?? string.Empty;
        
        _logger?.Info($"Сохраняем настройки оверлея: CRM='{crm}', Касса='{cash}', Display='{_overlayDisplay}'");
        
        try
        {
            var settingsDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Bobrus");
            var settingsPath = Path.Combine(settingsDir, "settings.json");
            
            AppSettings? existingSettings = null;
            if (File.Exists(settingsPath))
            {
                var json = await File.ReadAllTextAsync(settingsPath);
                existingSettings = JsonSerializer.Deserialize<AppSettings>(json);
            }
            
            var newSettings = new AppSettings(
                HideToTray: existingSettings?.HideToTray ?? true,
                Autostart: existingSettings?.Autostart ?? true,
                Theme: existingSettings?.Theme ?? "Dark",
                ShowAllSections: existingSettings?.ShowAllSections ?? false,
                ShowConsole: existingSettings?.ShowConsole ?? true,
                OverlayEnabled: true,
                OverlayCrm: crm,
                OverlayCashDesk: cash,
                OverlayDisplay: string.IsNullOrEmpty(_overlayDisplay) ? existingSettings?.OverlayDisplay : _overlayDisplay
            );
            
            Directory.CreateDirectory(settingsDir);
            var newJson = JsonSerializer.Serialize(newSettings, new JsonSerializerOptions { WriteIndented = true });
            await File.WriteAllTextAsync(settingsPath, newJson);
            
            _logger?.Info($"✔ Настройки оверлея сохранены: CRM={crm}, Касса={cash}");
            try
            {
                var exePath = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName;
                if (!string.IsNullOrEmpty(exePath))
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = exePath,
                        Arguments = "--overlay-runner",
                        UseShellExecute = false,
                        CreateNoWindow = true
                    });
                    _logger?.Info("✔ Overlay-runner запущен");
                }
            }
            catch (Exception runEx)
            {
                _logger?.Error($"Не удалось запустить overlay-runner: {runEx.Message}");
            }
        }
        catch (Exception ex)
        {
            _logger?.Error($"Ошибка сохранения настроек оверлея: {ex.Message}");
        }
    }

    private static bool IsStepCompleted(string line) => line.StartsWith("✔") || line.StartsWith("✖");

    private static bool TryParseStepStart(string line, int enabledSteps, out int current, out int total, out string title)
    {
        current = 0;
        total = enabledSteps;
        title = string.Empty;

        if (string.IsNullOrWhiteSpace(line) || !line.StartsWith("["))
        {
            return false;
        }

        var closingBracket = line.IndexOf(']');
        if (closingBracket <= 1)
        {
            return false;
        }

        var bracketContent = line.Substring(1, closingBracket - 1);
        var parts = bracketContent.Split('/');
        if (parts.Length != 2 ||
            !int.TryParse(parts[0], out current) ||
            !int.TryParse(parts[1], out var parsedTotal))
        {
            return false;
        }

        total = parsedTotal;
        title = line[(closingBracket + 1)..].Trim().TrimEnd('.');
        return current > 0 && total > 0;
    }

    private string FormatEta(int completed, int total)
    {
        if (completed <= 0 || total <= 0)
        {
            return "—";
        }

        var elapsed = DateTime.UtcNow - _setupStartUtc;
        var avgPerStep = TimeSpan.FromSeconds(elapsed.TotalSeconds / completed);
        var remainingSteps = Math.Max(0, total - completed);
        var eta = TimeSpan.FromSeconds(avgPerStep.TotalSeconds * remainingSteps);

        static string FormatSpan(TimeSpan ts)
        {
            if (ts.TotalHours >= 1)
            {
                return $"{(int)ts.TotalHours}ч {ts.Minutes:D2}м";
            }

            if (ts.TotalMinutes >= 1)
            {
                return $"{ts.Minutes}м {ts.Seconds:D2}с";
            }

            return $"{Math.Max(1, ts.Seconds)}с";
        }

        return FormatSpan(eta);
    }

    private void StartDragFromMaximized(MouseButtonEventArgs e)
    {
        var mousePos = e.GetPosition(this);
        var percentX = mousePos.X / ActualWidth;
        var percentY = mousePos.Y / ActualHeight;
        var screenPos = PointToScreen(mousePos);

        WindowState = WindowState.Normal;

        var targetWidth = _restoreWidth > 0 ? _restoreWidth : RestoreBounds.Width;
        var targetHeight = _restoreHeight > 0 ? _restoreHeight : RestoreBounds.Height;

        if (targetWidth > 0 && targetHeight > 0)
        {
            Width = targetWidth;
            Height = targetHeight;
        }

        Left = screenPos.X - (targetWidth * percentX);
        Top = screenPos.Y - (targetHeight * percentY);

        Dispatcher.BeginInvoke(new Action(() =>
        {
            try
            {
                DragMove();
            }
            catch
            {
            }
        }), DispatcherPriority.Input);
    }

    private void OnWindowLocationChanged(object? sender, EventArgs e) => CaptureRestoreBounds();

    private void OnWindowSizeChanged(object sender, SizeChangedEventArgs e) => CaptureRestoreBounds();

    private void CaptureRestoreBounds()
    {
        if (WindowState != WindowState.Normal)
        {
            return;
        }

        _restoreWidth = ActualWidth > 0 ? ActualWidth : Width;
        _restoreHeight = ActualHeight > 0 ? ActualHeight : Height;
        _restoreLeft = double.IsNaN(Left) ? RestoreBounds.Left : Left;
        _restoreTop = double.IsNaN(Top) ? RestoreBounds.Top : Top;
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == WmGetMinMaxInfo)
        {
            HandleMinMaxInfo(hwnd, lParam);
            handled = false;
        }

        return IntPtr.Zero;
    }

    private void HandleMinMaxInfo(IntPtr hwnd, IntPtr lParam)
    {
        var mmi = Marshal.PtrToStructure<MINMAXINFO>(lParam);
        var monitor = MonitorFromWindow(hwnd, MonitorDefaultToNearest);

        if (monitor != IntPtr.Zero)
        {
            var info = new MONITORINFO { cbSize = Marshal.SizeOf<MONITORINFO>() };
            if (GetMonitorInfo(monitor, ref info))
            {
                var work = info.rcWork;
                var monitorArea = info.rcMonitor;

                mmi.ptMaxPosition.X = work.Left - monitorArea.Left;
                mmi.ptMaxPosition.Y = work.Top - monitorArea.Top;
                mmi.ptMaxSize.X = work.Right - work.Left;
                mmi.ptMaxSize.Y = work.Bottom - work.Top;
            }
        }

        Marshal.StructureToPtr(mmi, lParam, true);
    }

    private const int MonitorDefaultToNearest = 0x00000002;
    private const int WmGetMinMaxInfo = 0x0024;

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MONITORINFO
    {
        public int cbSize;
        public RECT rcMonitor;
        public RECT rcWork;
        public int dwFlags;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT
    {
        public int X;
        public int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MINMAXINFO
    {
        public POINT ptReserved;
        public POINT ptMaxSize;
        public POINT ptMaxPosition;
        public POINT ptMinTrackSize;
        public POINT ptMaxTrackSize;
    }

    [DllImport("User32.dll")]
    private static extern IntPtr MonitorFromWindow(IntPtr hwnd, int dwFlags);

    [DllImport("User32.dll", SetLastError = true)]
    private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFO lpmi);

    private void AdjustResizeBorder()
    {
        var chrome = WindowChrome.GetWindowChrome(this);
        if (chrome is null)
        {
            chrome = new WindowChrome();
            WindowChrome.SetWindowChrome(this, chrome);
        }

        chrome.ResizeBorderThickness = WindowState == WindowState.Maximized
            ? new Thickness(0)
            : _defaultResizeBorder;
    }

    private void ShowCustomDialog(string title, string message, bool isError = false)
    {
        var dialog = new Window
        {
            Title = title,
            Width = 400,
            Height = 180,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Owner = this,
            WindowStyle = WindowStyle.None,
            ResizeMode = ResizeMode.NoResize,
            Background = (System.Windows.Media.Brush)FindResource("BackgroundBrush"),
            AllowsTransparency = true
        };

        var grid = new System.Windows.Controls.Grid { Margin = new Thickness(24) };
        grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });

        var titleText = new System.Windows.Controls.TextBlock
        {
            Text = title,
            FontSize = 18,
            FontWeight = FontWeights.SemiBold,
            Foreground = isError ? System.Windows.Media.Brushes.OrangeRed : (System.Windows.Media.Brush)FindResource("TextPrimaryBrush"),
            Margin = new Thickness(0, 0, 0, 12)
        };
        System.Windows.Controls.Grid.SetRow(titleText, 0);

        var messageText = new System.Windows.Controls.TextBlock
        {
            Text = message,
            TextWrapping = TextWrapping.Wrap,
            FontSize = 14,
            Foreground = (System.Windows.Media.Brush)FindResource("TextPrimaryBrush")
        };
        System.Windows.Controls.Grid.SetRow(messageText, 1);

        var okButton = new System.Windows.Controls.Button
        {
            Content = "OK",
            Width = 100,
            Height = 36,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Right,
            Style = (Style)FindResource("PrimaryButton")
        };
        okButton.Click += (s, e) => dialog.Close();
        System.Windows.Controls.Grid.SetRow(okButton, 2);

        grid.Children.Add(titleText);
        grid.Children.Add(messageText);
        grid.Children.Add(okButton);

        var border = new System.Windows.Controls.Border
        {
            Background = (System.Windows.Media.Brush)FindResource("SurfaceBrush"),
            BorderBrush = (System.Windows.Media.Brush)FindResource("BorderBrushMuted"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(12),
            Child = grid
        };

        dialog.Content = border;
        dialog.ShowDialog();
    }

    private bool ShowConfirmationDialog(string title, string message)
    {
        bool result = false;

        var dialog = new Window
        {
            Title = title,
            Width = 440,
            Height = 200,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Owner = this,
            WindowStyle = WindowStyle.None,
            ResizeMode = ResizeMode.NoResize,
            Background = System.Windows.Media.Brushes.Transparent,
            AllowsTransparency = true
        };

        var grid = new System.Windows.Controls.Grid { Margin = new Thickness(24) };
        grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });

        var titleText = new System.Windows.Controls.TextBlock
        {
            Text = title,
            FontSize = 18,
            FontWeight = FontWeights.SemiBold,
            Foreground = (System.Windows.Media.Brush)FindResource("TextPrimaryBrush"),
            Margin = new Thickness(0, 0, 0, 12)
        };
        System.Windows.Controls.Grid.SetRow(titleText, 0);

        var messageText = new System.Windows.Controls.TextBlock
        {
            Text = message,
            TextWrapping = TextWrapping.Wrap,
            FontSize = 14,
            Foreground = (System.Windows.Media.Brush)FindResource("TextPrimaryBrush")
        };
        System.Windows.Controls.Grid.SetRow(messageText, 1);

        var buttonPanel = new System.Windows.Controls.StackPanel
        {
            Orientation = System.Windows.Controls.Orientation.Horizontal,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Right,
            Margin = new Thickness(0, 16, 0, 0)
        };

        var cancelButton = new System.Windows.Controls.Button
        {
            Content = "Отмена",
            Width = 100,
            Height = 36,
            Margin = new Thickness(0, 0, 12, 0),
            Background = (System.Windows.Media.Brush)FindResource("SurfaceBrush"),
            Foreground = (System.Windows.Media.Brush)FindResource("TextPrimaryBrush"),
            BorderBrush = (System.Windows.Media.Brush)FindResource("BorderBrushMuted"),
            BorderThickness = new Thickness(1)
        };
        cancelButton.Click += (s, e) => { result = false; dialog.Close(); };

        var confirmButton = new System.Windows.Controls.Button
        {
            Content = "Начать",
            Width = 100,
            Height = 36,
            Style = (Style)FindResource("PrimaryButton")
        };
        confirmButton.Click += (s, e) => { result = true; dialog.Close(); };

        buttonPanel.Children.Add(cancelButton);
        buttonPanel.Children.Add(confirmButton);
        System.Windows.Controls.Grid.SetRow(buttonPanel, 2);

        grid.Children.Add(titleText);
        grid.Children.Add(messageText);
        grid.Children.Add(buttonPanel);

        var border = new System.Windows.Controls.Border
        {
            Background = (System.Windows.Media.Brush)FindResource("SurfaceBrush"),
            BorderBrush = (System.Windows.Media.Brush)FindResource("BorderBrushMuted"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(12),
            Child = grid
        };

        dialog.Content = border;
        dialog.ShowDialog();

        return result;
    }

}
