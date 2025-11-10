using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using PowerPresenter.Core.Interfaces;
using PowerPresenter.Core.Models;

namespace PowerPresenter.App.Services;

/// <summary>
/// Generates previews by leveraging the PowerPoint COM automation API.
/// </summary>
public sealed class InteropPreviewGenerationStrategy : IPreviewGenerationStrategy
{
    public string Name => "PowerPointInterop";

    public bool CanHandle(string presentationPath) => File.Exists(presentationPath);

    public async ValueTask<PreviewResult> GenerateAsync(string presentationPath, PreviewGenerationContext context, CancellationToken cancellationToken)
    {
        try
        {
            var cachePath = context.CachePathResolver(presentationPath);
            if (File.Exists(cachePath))
            {
                return PreviewResult.SuccessResult(cachePath);
            }

            return await Task.Run(() => GeneratePreviewInternal(presentationPath, cachePath), cancellationToken).ConfigureAwait(false);
        }
        catch (COMException ex)
        {
            return PreviewResult.Failure($"Interop non disponibile: {ex.Message}");
        }
        catch (Exception ex)
        {
            return PreviewResult.Failure(ex.Message);
        }
    }

    private static PreviewResult GeneratePreviewInternal(string presentationPath, string cachePath)
    {
        dynamic? application = null;
        dynamic? presentation = null;
        try
        {
            var type = Type.GetTypeFromProgID("PowerPoint.Application");
            if (type is null)
            {
                return PreviewResult.Failure("PowerPoint non è installato.");
            }

            application = Activator.CreateInstance(type);
            if (application is null)
            {
                return PreviewResult.Failure("Impossibile inizializzare PowerPoint.");
            }

            application.Visible = -1; // msoTrue

            presentation = application.Presentations.Open(
                presentationPath,
                WithWindow: 0,   // msoFalse
                ReadOnly: -1,    // msoTrue
                Untitled: 0);    // msoFalse

            var slide = presentation.Slides[1];
            var tempPath = Path.ChangeExtension(Path.GetTempFileName(), ".png");
            slide.Export(tempPath, "PNG", 960, 540);
            Directory.CreateDirectory(Path.GetDirectoryName(cachePath)!);
            File.Copy(tempPath, cachePath, true);
            File.Delete(tempPath);
            return PreviewResult.SuccessResult(cachePath);
        }
        finally
        {
            if (presentation is not null)
            {
                try
                {
                    presentation.Close();
                }
                catch
                {
                    // ignore cleanup issues
                }
                ReleaseComObject(presentation);
            }

            if (application is not null)
            {
                ReleaseComObject(application);
            }
        }
    }

    private static void ReleaseComObject(object comObject)
    {
        try
        {
            Marshal.ReleaseComObject(comObject);
        }
        catch
        {
            // ignored
        }
    }
}
