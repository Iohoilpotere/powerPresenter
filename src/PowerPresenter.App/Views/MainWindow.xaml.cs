using System.Windows;
using PowerPresenter.App.ViewModels;

namespace PowerPresenter.App.Views;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }

    public MainWindowViewModel ViewModel => (MainWindowViewModel)DataContext;

    private void OnMonitorButtonClick(object sender, RoutedEventArgs e)
    {
        if (sender is System.Windows.Controls.Button button && button.ContextMenu is { } menu)
        {
            menu.PlacementTarget = button;
            menu.IsOpen = true;
        }
    }
}
