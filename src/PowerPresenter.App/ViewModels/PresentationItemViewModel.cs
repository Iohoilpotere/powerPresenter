using PowerPresenter.Core.Models;

namespace PowerPresenter.App.ViewModels;

public sealed class PresentationItemViewModel
{
    public PresentationItemViewModel(PresentationItem item)
    {
        Item = item;
        DisplayName = item.DisplayName;
        FilePath = item.FilePath;
    }

    public PresentationItem Item { get; }

    public string DisplayName { get; }

    public string FilePath { get; }

    public string? PreviewPath { get; private set; }

    public void SetPreview(string? previewPath)
    {
        PreviewPath = previewPath;
    }
}
