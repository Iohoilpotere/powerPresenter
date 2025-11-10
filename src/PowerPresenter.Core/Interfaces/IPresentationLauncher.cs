using PowerPresenter.Core.Enums;

namespace PowerPresenter.Core.Interfaces;

/// <summary>
/// Launches PowerPoint presentations.
/// </summary>
public interface IPresentationLauncher
{
    Task LaunchAsync(string presentationPath, MonitorPreference preference, CancellationToken cancellationToken);
}
