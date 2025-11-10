using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Interop.PowerPoint;
using PowerPresenter.Core.Enums;
using PowerPresenter.Core.Interfaces;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

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

            var slideShowSettings = presentation.SlideShowSettings;
            slideShowSettings.LoopUntilStopped = Office.MsoTriState.msoFalse;
            slideShowSettings.ShowWithAnimation = Office.MsoTriState.msoTrue;
            slideShowSettings.RangeType = PowerPoint.PpSlideShowRangeType.ppShowAll;

            slideShowSettings.ShowPresenterView = preference == MonitorPreference.Secondary
                ? Office.MsoTriState.msoTrue
                : Office.MsoTriState.msoFalse;
            slideShowSettings.Run();
        }
        finally
        {
            if (presentation is not null)
            {
                Marshal.ReleaseComObject(presentation);
            }

            if (application is not null)
            {
                Marshal.ReleaseComObject(application);
            }
        }
    }
}
