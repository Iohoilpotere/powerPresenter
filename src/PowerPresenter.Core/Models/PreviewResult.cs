namespace PowerPresenter.Core.Models;

/// <summary>
/// Result of a preview generation attempt.
/// </summary>
public sealed class PreviewResult
{
    private PreviewResult(bool success, string? previewImagePath, string? error)
    {
        Success = success;
        PreviewImagePath = previewImagePath;
        Error = error;
    }

    public bool Success { get; }

    public string? PreviewImagePath { get; }

    public string? Error { get; }

    public static PreviewResult Failure(string error) => new(false, null, error);

    public static PreviewResult SuccessResult(string previewImagePath) => new(true, previewImagePath, null);
}
