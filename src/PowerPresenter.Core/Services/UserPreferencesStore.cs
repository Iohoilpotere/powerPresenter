using System;
using System.IO;
using System.Text.Json;
using PowerPresenter.Core.Configuration;
using PowerPresenter.Core.Interfaces;

namespace PowerPresenter.Core.Services;

/// <summary>
/// Thread-safe singleton responsible for persisting user preferences.
/// </summary>
public sealed class UserPreferencesStore : IUserPreferencesStore
{
    private static readonly Lazy<UserPreferencesStore> LazyInstance = new(() => new UserPreferencesStore());
    private readonly string _storagePath;

    private UserPreferencesStore()
    {
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var directory = Path.Combine(appData, "PowerPresenter");
        Directory.CreateDirectory(directory);
        _storagePath = Path.Combine(directory, "user-preferences.json");
        Current = Load();
    }

    public static UserPreferencesStore Instance => LazyInstance.Value;

    public UserPreferences Current { get; }

    public Task SaveAsync(CancellationToken cancellationToken)
    {
        var json = JsonSerializer.Serialize(Current, new JsonSerializerOptions { WriteIndented = true });
        return File.WriteAllTextAsync(_storagePath, json, cancellationToken);
    }

    private UserPreferences Load()
    {
        if (!File.Exists(_storagePath))
        {
            return new UserPreferences();
        }

        try
        {
            var json = File.ReadAllText(_storagePath);
            return JsonSerializer.Deserialize<UserPreferences>(json) ?? new UserPreferences();
        }
        catch
        {
            return new UserPreferences();
        }
    }
}
