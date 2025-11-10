using PowerPresenter.Core.Configuration;

namespace PowerPresenter.Core.Interfaces;

/// <summary>
/// Provides access to persisted user preferences.
/// </summary>
public interface IUserPreferencesStore
{
    UserPreferences Current { get; }

    Task SaveAsync(CancellationToken cancellationToken);
}
