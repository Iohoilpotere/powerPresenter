using System.IO;
using System.Security.Cryptography;
using System.Text;
using PowerPresenter.Core.Interfaces;
using PowerPresenter.Core.Models;

namespace PowerPresenter.Core.Services;

/// <summary>
/// Provides a deterministic cache for preview images.
/// </summary>
public sealed class PreviewCacheService : IPreviewCacheService
{
    private readonly string _cacheDirectory;

    public PreviewCacheService(string cacheDirectory)
    {
        _cacheDirectory = cacheDirectory;
        Directory.CreateDirectory(_cacheDirectory);
    }

    public PreviewGenerationContext CreateContext() => new(
        _cacheDirectory,
        GetCachePathFor);

    public string GetCachePathFor(string presentationPath)
    {
        using var sha = SHA256.Create();
        var hash = sha.ComputeHash(Encoding.UTF8.GetBytes(presentationPath));
        var fileName = Convert.ToHexString(hash) + ".png";
        return Path.Combine(_cacheDirectory, fileName);
    }
}
