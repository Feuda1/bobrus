namespace Bobrus.App;

internal sealed record AppSettings(
    bool HideToTray,
    bool Autostart,
    string? Theme,
    bool ShowAllSections,
    bool ShowConsole,
    bool OverlayEnabled,
    string? OverlayCrm,
    string? OverlayCashDesk,
    string? OverlayDisplay);
