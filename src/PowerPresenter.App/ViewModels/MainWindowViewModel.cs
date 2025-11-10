using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using PowerPresenter.App.Commands;
using PowerPresenter.Core.Configuration;
using PowerPresenter.Core.Enums;
using PowerPresenter.Core.Interfaces;
using PowerPresenter.Core.Models;

namespace PowerPresenter.App.ViewModels;

public sealed class MainWindowViewModel : INotifyPropertyChanged
{
    private readonly IPresentationDiscoveryService _discoveryService;
    private readonly IPreviewGenerationStrategyFactory _strategyFactory;
    private readonly IPreviewCacheService _previewCacheService;
    private readonly IPresentationLauncher _launcher;
    private readonly IUserPreferencesStore _preferencesStore;
    private readonly IMonitorService _monitorService;
    private readonly PreviewGenerationContext _previewContext;
    private readonly Brush _defaultBackground;
    private CancellationTokenSource? _loadingCancellationTokenSource;
    private string? _currentFolder;
    private bool _isBusy;
    private string _statusMessage = string.Empty;
    private Brush? _backgroundBrush;

    public MainWindowViewModel(
        IPresentationDiscoveryService discoveryService,
        IPreviewGenerationStrategyFactory strategyFactory,
        IPreviewCacheService previewCacheService,
        IPresentationLauncher launcher,
        IUserPreferencesStore preferencesStore,
        IMonitorService monitorService,
        PreviewGenerationContext previewContext)
    {
        _discoveryService = discoveryService;
        _strategyFactory = strategyFactory;
        _previewCacheService = previewCacheService;
        _launcher = launcher;
        _preferencesStore = preferencesStore;
        _monitorService = monitorService;
        _previewContext = previewContext;
        _defaultBackground = new SolidColorBrush(Color.FromRgb(16, 21, 34));
        _defaultBackground.Freeze();

        Presentations = new ObservableCollection<PresentationItemViewModel>();
        Presentations.CollectionChanged += (_, _) => OnPropertyChanged(nameof(IsEmpty));

        RefreshCommand = new AsyncRelayCommand(_ => RefreshAsync(), _ => !string.IsNullOrWhiteSpace(_currentFolder));
        BrowseFolderCommand = new RelayCommand(BrowseFolder);
        CloseCommand = new RelayCommand(() => Application.Current?.Shutdown());
        LaunchPresentationCommand = new AsyncRelayCommand(LaunchPresentationAsync, parameter => parameter is PresentationItemViewModel);
        SetBackgroundCommand = new RelayCommand(SetBackground);
        ChangeMonitorPreferenceCommand = new AsyncRelayCommand(ChangeMonitorPreferenceAsync);

        ApplyPreferences(_preferencesStore.Current);
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public ObservableCollection<PresentationItemViewModel> Presentations { get; }

    public AsyncRelayCommand RefreshCommand { get; }

    public RelayCommand BrowseFolderCommand { get; }

    public RelayCommand CloseCommand { get; }

    public AsyncRelayCommand LaunchPresentationCommand { get; }

    public RelayCommand SetBackgroundCommand { get; }

    public AsyncRelayCommand ChangeMonitorPreferenceCommand { get; }

    public bool IsBusy
    {
        get => _isBusy;
        private set
        {
            if (_isBusy != value)
            {
                _isBusy = value;
                OnPropertyChanged(nameof(IsBusy));
                OnPropertyChanged(nameof(IsEmpty));
            }
        }
    }

    public bool IsEmpty => Presentations.Count == 0 && !IsBusy;

    public string? CurrentFolder
    {
        get => _currentFolder;
        private set
        {
            if (_currentFolder != value)
            {
                _currentFolder = value;
                OnPropertyChanged(nameof(CurrentFolder));
                RefreshCommand.RaiseCanExecuteChanged();
            }
        }
    }

    public string StatusMessage
    {
        get => _statusMessage;
        private set
        {
            if (_statusMessage != value)
            {
                _statusMessage = value;
                OnPropertyChanged(nameof(StatusMessage));
            }
        }
    }

    public Brush? BackgroundBrush
    {
        get => _backgroundBrush;
        private set
        {
            if (_backgroundBrush != value)
            {
                _backgroundBrush = value;
                OnPropertyChanged(nameof(BackgroundBrush));
            }
        }
    }

    public string MonitorPreferenceLabel => _preferencesStore.Current.MonitorPreference switch
    {
        MonitorPreference.Automatic => "Automatico",
        MonitorPreference.Primary => "Primario",
        MonitorPreference.Secondary => "Esteso",
        _ => "Automatico"
    };

    public void SetInitialFolder(string folderPath)
    {
        if (Directory.Exists(folderPath))
        {
            CurrentFolder = folderPath;
            _ = RefreshAsync();
        }
        else
        {
            StatusMessage = "La cartella fornita non esiste.";
        }
    }

    private void BrowseFolder()
    {
        using var dialog = new FolderBrowserDialog
        {
            Description = "Seleziona la cartella contenente le presentazioni",
            UseDescriptionForTitle = true
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            CurrentFolder = dialog.SelectedPath;
            _ = RefreshAsync();
        }
    }

    private async Task RefreshAsync()
    {
        if (string.IsNullOrWhiteSpace(_currentFolder))
        {
            StatusMessage = "Selezionare una cartella";
            return;
        }

        if (!Directory.Exists(_currentFolder))
        {
            StatusMessage = "La cartella non esiste più";
            return;
        }

        _loadingCancellationTokenSource?.Cancel();
        _loadingCancellationTokenSource = new CancellationTokenSource();
        var token = _loadingCancellationTokenSource.Token;

        try
        {
            IsBusy = true;
            StatusMessage = "Caricamento presentazioni...";

            var items = await Task.Run(() => _discoveryService.Discover(_currentFolder!).ToList(), token).ConfigureAwait(false);

            await Application.Current!.Dispatcher.InvokeAsync(() =>
            {
                Presentations.Clear();
                foreach (var item in items)
                {
                    Presentations.Add(new PresentationItemViewModel(item));
                }
            });

            if (Presentations.Count == 0)
            {
                StatusMessage = "Nessun file PowerPoint trovato.";
                return;
            }

            StatusMessage = $"Trovate {Presentations.Count} presentazioni.";

            var strategies = _strategyFactory.BuildStrategies();
            var previewTasks = Presentations.Select(vm => Task.Run(async () =>
            {
                foreach (var strategy in strategies)
                {
                    if (!strategy.CanHandle(vm.FilePath))
                    {
                        continue;
                    }

                    var result = await strategy.GenerateAsync(vm.FilePath, _previewContext, token).ConfigureAwait(false);
                    if (result.Success && !string.IsNullOrWhiteSpace(result.PreviewImagePath))
                    {
                        return (vm, result.PreviewImagePath);
                    }
                }

                return (vm, (string?)null);
            }, token)).ToList();

            var previews = await Task.WhenAll(previewTasks).ConfigureAwait(false);

            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                foreach (var (vm, path) in previews)
                {
                    vm.SetPreview(path);
                }
            });
        }
        catch (OperationCanceledException)
        {
            // intentionally ignored - user triggered refresh cancellation
        }
        finally
        {
            IsBusy = false;
        }
    }

    private async Task LaunchPresentationAsync(object? parameter)
    {
        if (parameter is not PresentationItemViewModel viewModel)
        {
            return;
        }

        try
        {
            await _launcher.LaunchAsync(viewModel.FilePath, _preferencesStore.Current.MonitorPreference, CancellationToken.None);
        }
        catch (Exception ex)
        {
            StatusMessage = $"Errore durante l'apertura: {ex.Message}";
        }
    }

    private void SetBackground()
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "Immagini (*.png;*.jpg;*.jpeg)|*.png;*.jpg;*.jpeg|Tutti i file (*.*)|*.*",
            Title = "Seleziona uno sfondo"
        };

        if (dialog.ShowDialog() == true)
        {
            ApplyBackground(dialog.FileName);
            _preferencesStore.Current.BackgroundImagePath = dialog.FileName;
            _ = _preferencesStore.SaveAsync(CancellationToken.None);
        }
    }

    private async Task ChangeMonitorPreferenceAsync(object? parameter)
    {
        if (parameter is not MonitorPreference preference)
        {
            return;
        }

        _preferencesStore.Current.MonitorPreference = preference;
        await _preferencesStore.SaveAsync(CancellationToken.None).ConfigureAwait(false);
        OnPropertyChanged(nameof(MonitorPreferenceLabel));
        StatusMessage = preference switch
        {
            MonitorPreference.Automatic => "Modalità monitor: automatico",
            MonitorPreference.Primary => "Modalità monitor: primario",
            MonitorPreference.Secondary => "Modalità monitor: esteso",
            _ => StatusMessage
        };
    }

    private void ApplyPreferences(UserPreferences preferences)
    {
        ApplyBackground(preferences.BackgroundImagePath);
        OnPropertyChanged(nameof(MonitorPreferenceLabel));

        if (_monitorService.HasMultipleMonitors())
        {
            StatusMessage = "Più monitor rilevati";
        }
        else
        {
            StatusMessage = "Utilizzo del monitor corrente";
        }
    }

    private void ApplyBackground(string? imagePath)
    {
        if (string.IsNullOrWhiteSpace(imagePath) || !File.Exists(imagePath))
        {
            BackgroundBrush = _defaultBackground;
            return;
        }

        try
        {
            var image = new BitmapImage();
            image.BeginInit();
            image.UriSource = new Uri(imagePath, UriKind.Absolute);
            image.CacheOption = BitmapCacheOption.OnLoad;
            image.EndInit();
            image.Freeze();

            var brush = new ImageBrush(image)
            {
                Stretch = Stretch.UniformToFill
            };
            brush.Freeze();
            BackgroundBrush = brush;
        }
        catch
        {
            BackgroundBrush = _defaultBackground;
        }
    }

    private void OnPropertyChanged(string propertyName)
    {
        var dispatcher = Application.Current?.Dispatcher;
        if (dispatcher is null || dispatcher.CheckAccess())
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        else
        {
            dispatcher.Invoke(() => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName)));
        }
    }
}
