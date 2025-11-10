using PowerPresenter.Core.Models;

namespace PowerPresenter.Core.Interfaces;

/// <summary>
/// Caches preview artefacts on disk.
/// </summary>
public interface IPreviewCacheService
{
    PreviewGenerationContext CreateContext();

    string GetCachePathFor(string presentationPath);
}
