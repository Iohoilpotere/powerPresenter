namespace PowerPresenter.Core.Models;

/// <summary>
/// Describes a PowerPoint presentation available for playback.
/// </summary>
public sealed class PresentationItem
{
    public PresentationItem(string filePath, string displayName)
    {
        FilePath = filePath;
        DisplayName = displayName;
    }

    /// <summary>
    /// Absolute path to the PowerPoint file.
    /// </summary>
    public string FilePath { get; }

    /// <summary>
    /// Display name presented to the operator.
    /// </summary>
    public string DisplayName { get; }
}
