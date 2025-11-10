namespace PowerPresenter.Core.Models;

/// <summary>
/// Encapsulates contextual data required to produce previews.
/// </summary>
public sealed class PreviewGenerationContext
{
    public PreviewGenerationContext(string cacheDirectory, Func<string, string> cachePathResolver)
    {
        CacheDirectory = cacheDirectory;
        CachePathResolver = cachePathResolver;
    }

    public string CacheDirectory { get; }

    public Func<string, string> CachePathResolver { get; }
}
