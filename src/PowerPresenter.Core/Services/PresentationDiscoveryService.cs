using System.IO;
using PowerPresenter.Core.Interfaces;
using PowerPresenter.Core.Models;

namespace PowerPresenter.Core.Services;

/// <summary>
/// Enumerates supported PowerPoint presentations within a directory.
/// </summary>
public sealed class PresentationDiscoveryService : IPresentationDiscoveryService
{
    private static readonly string[] SupportedExtensions = { ".pptx", ".ppt", ".ppsx", ".pps" };

    public IEnumerable<PresentationItem> Discover(string folderPath)
    {
        if (!Directory.Exists(folderPath))
        {
            return Enumerable.Empty<PresentationItem>();
        }

        return Directory.EnumerateFiles(folderPath)
            .Where(file => SupportedExtensions.Any(ext => file.EndsWith(ext, StringComparison.OrdinalIgnoreCase)))
            .Select(file => new PresentationItem(file, Path.GetFileNameWithoutExtension(file)))
            .OrderBy(item => item.DisplayName, StringComparer.CurrentCultureIgnoreCase)
            .ToList();
    }
}
