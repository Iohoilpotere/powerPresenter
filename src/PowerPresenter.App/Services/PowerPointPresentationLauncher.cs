using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using PowerPresenter.Core.Enums;
using PowerPresenter.Core.Interfaces;

namespace PowerPresenter.App.Services;

/// <summary>
/// Launches presentations through the PowerPoint COM automation layer.
/// </summary>
public sealed class PowerPointPresentationLauncher : IPresentationLauncher
{
    public async Task LaunchAsync(string presentationPath, MonitorPreference preference, CancellationToken cancellationToken)
    {
        await Task.Run(() => LaunchInternal(presentationPath, preference), cancellationToken).ConfigureAwait(false);
    }

    private void LaunchInternal(string presentationPath, MonitorPreference preference)
    {
        dynamic? application = null;
        dynamic? presentation = null;
        try
        {
            var type = Type.GetTypeFromProgID("PowerPoint.Application");
            if (type is null)
            {
                throw new InvalidOperationException("PowerPoint non è installato sul sistema.");
            }

            application = Activator.CreateInstance(type) ?? throw new InvalidOperationException("Impossibile inizializzare PowerPoint.");
            application.Visible = -1; // msoTrue

            presentation = application.Presentations.Open(
                presentationPath,
                WithWindow: 0,   // msoFalse
                ReadOnly: -1,    // msoTrue
                Untitled: 0);    // msoFalse

            var slideShowSettings = presentation.SlideShowSettings;
            slideShowSettings.LoopUntilStopped = 0; // msoFalse
            slideShowSettings.ShowWithAnimation = -1; // msoTrue
            slideShowSettings.RangeType = 1; // ppShowAll

            slideShowSettings.ShowPresenterView = preference == MonitorPreference.Secondary
                ? -1
                : 0;
            slideShowSettings.Run();
        }
        finally
        {
            if (presentation is not null)
            {
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
            // ignored on shutdown
        }
    }
}
