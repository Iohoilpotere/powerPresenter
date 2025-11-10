using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Interop.PowerPoint;
using PowerPresenter.Core.Interfaces;
using PowerPresenter.Core.Models;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

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
        PowerPoint.Application? application = null;
        PowerPoint.Presentation? presentation = null;
        try
        {
            application = new PowerPoint.Application
            {
                Visible = Office.MsoTriState.msoTrue
            };

            presentation = application.Presentations.Open(
                presentationPath,
                WithWindow: Office.MsoTriState.msoFalse,
                ReadOnly: Office.MsoTriState.msoTrue,
                Untitled: Office.MsoTriState.msoFalse);

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
                presentation.Close();
                Marshal.ReleaseComObject(presentation);
            }

            if (application is not null)
            {
                Marshal.ReleaseComObject(application);
            }
        }
    }
}
