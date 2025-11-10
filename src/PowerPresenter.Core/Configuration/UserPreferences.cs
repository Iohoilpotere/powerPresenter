using PowerPresenter.Core.Enums;

namespace PowerPresenter.Core.Configuration;

/// <summary>
/// Stores the persisted user-facing configuration.
/// </summary>
public sealed class UserPreferences
{
    public string? BackgroundImagePath { get; set; }

    public MonitorPreference MonitorPreference { get; set; } = MonitorPreference.Automatic;
}
