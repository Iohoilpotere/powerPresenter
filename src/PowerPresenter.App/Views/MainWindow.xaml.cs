using System.Windows;
using System.Windows.Controls;
using PowerPresenter.App.ViewModels;

namespace PowerPresenter.Views;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }

    public MainWindowViewModel ViewModel => (MainWindowViewModel)DataContext;

    private void OnMonitorButtonClick(object sender, RoutedEventArgs e)
    {
        if (sender is Button button && button.ContextMenu is { } menu)
        {
            menu.PlacementTarget = button;
            menu.IsOpen = true;
        }
    }
}
