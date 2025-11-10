using PowerPresenter.Core.Interfaces;

namespace PowerPresenter.Core.Services;

/// <summary>
/// Provides a simple composition over registered preview strategies.
/// </summary>
public sealed class PreviewGenerationStrategyFactory : IPreviewGenerationStrategyFactory
{
    private readonly IReadOnlyCollection<IPreviewGenerationStrategy> _strategies;

    public PreviewGenerationStrategyFactory(IEnumerable<IPreviewGenerationStrategy> strategies)
    {
        _strategies = strategies.ToList();
    }

    public IReadOnlyCollection<IPreviewGenerationStrategy> BuildStrategies() => _strategies;
}
