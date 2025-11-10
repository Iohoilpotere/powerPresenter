using PowerPresenter.Core.Models;

namespace PowerPresenter.Core.Interfaces;

/// <summary>
/// Discovers PowerPoint files within a folder.
/// </summary>
public interface IPresentationDiscoveryService
{
    IEnumerable<PresentationItem> Discover(string folderPath);
}
