using System.Threading;
using System.Threading.Tasks;
using PowerPresenter.Core.Enums;
using PowerPresenter.Core.Interfaces;

namespace PowerPresenter.App.Services;

/// <summary>
/// Decorator that normalises monitor selection requests before launching a presentation.
/// </summary>
public sealed class MonitorAwarePresentationLauncherDecorator : IPresentationLauncher
{
    private readonly IPresentationLauncher _inner;
    private readonly IMonitorService _monitorService;

    public MonitorAwarePresentationLauncherDecorator(IPresentationLauncher inner, IMonitorService monitorService)
    {
        _inner = inner;
        _monitorService = monitorService;
    }

    public Task LaunchAsync(string presentationPath, MonitorPreference preference, CancellationToken cancellationToken)
    {
        var resolvedPreference = preference switch
        {
            MonitorPreference.Automatic => _monitorService.HasMultipleMonitors() ? MonitorPreference.Secondary : MonitorPreference.Primary,
            _ => preference
        };

        return _inner.LaunchAsync(presentationPath, resolvedPreference, cancellationToken);
    }
}
