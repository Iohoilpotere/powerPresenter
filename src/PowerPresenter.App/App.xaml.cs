using System;
using System.IO;
using System.Windows;
using PowerPresenter.Core.Interfaces;
using PowerPresenter.Core.Services;
using PowerPresenter.App.Services;
using PowerPresenter.App.ViewModels;

namespace PowerPresenter;

public partial class App : Application
{
    private IPreviewCacheService? _previewCacheService;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        var cacheDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PowerPresenter", "Previews");
        _previewCacheService = new PreviewCacheService(cacheDirectory);

        var previewContext = _previewCacheService.CreateContext();

        var thumbnailStrategy = new ThumbnailPreviewGenerationStrategy();
        var interopStrategy = new InteropPreviewGenerationStrategy();
        var strategyFactory = new PreviewGenerationStrategyFactory(new IPreviewGenerationStrategy[]
        {
            interopStrategy,
            thumbnailStrategy
        });

        var presentationDiscoveryService = new PresentationDiscoveryService();
        var monitorService = new MonitorService();
        var preferencesStore = UserPreferencesStore.Instance;
        var launcher = new MonitorAwarePresentationLauncherDecorator(new PowerPointPresentationLauncher(), monitorService);

        var mainWindow = new Views.MainWindow
        {
            DataContext = new MainWindowViewModel(
                presentationDiscoveryService,
                strategyFactory,
                _previewCacheService,
                launcher,
                preferencesStore,
                monitorService,
                previewContext)
        };

        if (e.Args.Length > 0)
        {
            mainWindow.ViewModel.SetInitialFolder(e.Args[0]);
        }

        mainWindow.Show();
    }
}
