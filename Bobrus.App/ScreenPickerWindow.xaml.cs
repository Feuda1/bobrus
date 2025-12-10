using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using WinForms = System.Windows.Forms;

namespace Bobrus.App;

public sealed partial class ScreenPickerWindow : Window
{
    public string? SelectedDeviceName { get; private set; }
    private readonly string? _currentDevice;

    public ScreenPickerWindow(string? currentDevice)
    {
        InitializeComponent();
        _currentDevice = currentDevice;
        Loaded += OnLoaded;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        var screens = WinForms.Screen.AllScreens;
        var items = screens.Select((s, index) => new ScreenItem
        {
            DeviceName = s.DeviceName,
            Title = $"{index + 1}: {(s.Primary ? "Основной" : "Экран")}",
            Resolution = $"{s.Bounds.Width}x{s.Bounds.Height}",
            PrimaryLabel = s.Primary ? "Текущий основной" : "Вторичный"
        }).ToList();

        ScreenList.ItemsSource = items;

        var selected = items.FirstOrDefault(i =>
            string.Equals(i.DeviceName, _currentDevice, StringComparison.OrdinalIgnoreCase));

        ScreenList.SelectedItem = selected ?? items.FirstOrDefault();
    }

    private void OnSelect(object sender, RoutedEventArgs e)
    {
        if (ScreenList.SelectedItem is ScreenItem item)
        {
            SelectedDeviceName = item.DeviceName;
            DialogResult = true;
        }
        else
        {
            DialogResult = false;
        }
        Close();
    }

    private void OnCancel(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    private void OnScreenDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        OnSelect(sender, e);
    }

    private sealed record ScreenItem
    {
        public string DeviceName { get; init; } = string.Empty;
        public string Title { get; init; } = string.Empty;
        public string Resolution { get; init; } = string.Empty;
        public string PrimaryLabel { get; init; } = string.Empty;
    }
}
