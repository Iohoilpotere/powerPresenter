using PowerPresenter.Core.Interfaces;

namespace PowerPresenter.App.Services;

public sealed class MonitorService : IMonitorService
{
    public bool HasMultipleMonitors() => System.Windows.Forms.Screen.AllScreens.Length > 1;
}
