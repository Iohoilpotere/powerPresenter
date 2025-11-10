using PowerPresenter.Core.Models;

namespace PowerPresenter.Core.Interfaces;

/// <summary>
/// Generates preview images for presentations.
/// </summary>
public interface IPreviewGenerationStrategy
{
    string Name { get; }

    ValueTask<PreviewResult> GenerateAsync(string presentationPath, PreviewGenerationContext context, CancellationToken cancellationToken);

    bool CanHandle(string presentationPath);
}
