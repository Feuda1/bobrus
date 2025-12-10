using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace Bobrus.App;

public partial class OverlayWindow : Window
{
    private const int GwlExstyle = -20;
    private const int WsExTransparent = 0x00000020;
    private const int WsExToolWindow = 0x00000080;
    private const int WsExLayered = 0x00080000;
    private Rect _targetBounds;

    public OverlayWindow()
    {
        InitializeComponent();
        SystemParameters.StaticPropertyChanged += OnSystemParametersChanged;
        Loaded += (_, _) => RefreshLayout();
        Closed += (_, _) => SystemParameters.StaticPropertyChanged -= OnSystemParametersChanged;
    }

    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);
        MakeClickThrough();
        RefreshLayout();
    }

    public void UpdateOverlay(string crm, string cashDesk, Rect bounds)
    {
        _targetBounds = bounds;
        OverlayText.Text = $"CRM: {crm}    Касса: {cashDesk}";
        OverlayText.FontSize = CalculateFontSize();
    }

    public void RefreshLayout()
    {
        if (_targetBounds.Width <= 0 || _targetBounds.Height <= 0)
        {
            _targetBounds = new Rect(SystemParameters.VirtualScreenLeft, SystemParameters.VirtualScreenTop, SystemParameters.VirtualScreenWidth, SystemParameters.VirtualScreenHeight);
        }

        Left = _targetBounds.Left;
        Top = _targetBounds.Top;
        Width = _targetBounds.Width;
        Height = _targetBounds.Height;
        OverlayText.FontSize = CalculateFontSize();
    }

    private double CalculateFontSize()
    {
        var referenceSize = Math.Min(_targetBounds.Width, _targetBounds.Height);
        if (referenceSize <= 0)
        {
            referenceSize = Math.Min(SystemParameters.VirtualScreenWidth, SystemParameters.VirtualScreenHeight);
        }

        return Math.Max(18, Math.Min(54, Math.Round(referenceSize * 0.03, 1)));
    }

    private void OnSystemParametersChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName is nameof(SystemParameters.VirtualScreenWidth) or nameof(SystemParameters.VirtualScreenHeight) or nameof(SystemParameters.WorkArea))
        {
            RefreshLayout();
        }
    }

    private void MakeClickThrough()
    {
        var helper = new WindowInteropHelper(this);
        var handle = helper.Handle;
        if (handle == IntPtr.Zero) return;

        var style = GetWindowLong(handle, GwlExstyle);
        SetWindowLong(handle, GwlExstyle, style | WsExTransparent | WsExToolWindow | WsExLayered);
    }

    [DllImport("user32.dll")]
    private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
}
