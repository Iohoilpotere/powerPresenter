namespace PowerPresenter.Core.Interfaces;

/// <summary>
/// Provides a composed preview strategy sequence.
/// </summary>
public interface IPreviewGenerationStrategyFactory
{
    IReadOnlyCollection<IPreviewGenerationStrategy> BuildStrategies();
}
