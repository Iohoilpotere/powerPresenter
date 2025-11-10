using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PowerPresenter.Core.Interfaces;
using PowerPresenter.Core.Models;
using A = DocumentFormat.OpenXml.Drawing;

namespace PowerPresenter.Core.Services;

/// <summary>
/// Attempts to reuse embedded thumbnails from the presentation package.
/// Falls back to generating a branded placeholder.
/// </summary>
public sealed class ThumbnailPreviewGenerationStrategy : IPreviewGenerationStrategy
{
    public string Name => "EmbeddedThumbnail";

    public bool CanHandle(string presentationPath) => Path.GetExtension(presentationPath).Equals(".pptx", StringComparison.OrdinalIgnoreCase);

    public async ValueTask<PreviewResult> GenerateAsync(string presentationPath, PreviewGenerationContext context, CancellationToken cancellationToken)
    {
        try
        {
            var cachePath = context.CachePathResolver(presentationPath);
            if (File.Exists(cachePath))
            {
                return PreviewResult.SuccessResult(cachePath);
            }

            await using var stream = File.OpenRead(presentationPath);
            using var presentation = PresentationDocument.Open(stream, false);
            var thumbnail = presentation.PresentationPart?.GetPartsOfType<ThumbnailPart>().FirstOrDefault();
            if (thumbnail is not null)
            {
                await using var thumbStream = thumbnail.GetStream(FileMode.Open, FileAccess.Read);
                Directory.CreateDirectory(Path.GetDirectoryName(cachePath)!);
                await using var target = File.Create(cachePath);
                await thumbStream.CopyToAsync(target, cancellationToken);
                return PreviewResult.SuccessResult(cachePath);
            }

            var title = ExtractFirstSlideTitle(presentation.PresentationPart);
            return CreatePlaceholder(cachePath, Path.GetFileNameWithoutExtension(presentationPath), title);
        }
        catch (Exception ex)
        {
            return PreviewResult.Failure(ex.Message);
        }
    }

    private static string? ExtractFirstSlideTitle(PresentationPart? presentationPart)
    {
        if (presentationPart?.Presentation?.SlideIdList is not SlideIdList slideIdList)
        {
            return null;
        }

        var firstSlideId = slideIdList.ChildElements.OfType<SlideId>().FirstOrDefault();
        if (firstSlideId is null)
        {
            return null;
        }

        var slidePart = presentationPart.GetPartById(firstSlideId.RelationshipId) as SlidePart;
        if (slidePart is null)
        {
            return null;
        }

        var text = slidePart.Slide.Descendants<A.Text>().Select(t => t.Text).FirstOrDefault(t => !string.IsNullOrWhiteSpace(t));
        return text;
    }

    private static PreviewResult CreatePlaceholder(string cachePath, string fileName, string? slideTitle)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(cachePath)!);
        const int width = 512;
        const int height = 288;
        using var bitmap = new Bitmap(width, height);
        using (var graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.FromArgb(24, 30, 54));
            using var titleFont = new Font("Segoe UI", 18, FontStyle.Bold, GraphicsUnit.Point);
            using var subtitleFont = new Font("Segoe UI", 12, FontStyle.Regular, GraphicsUnit.Point);
            using var accentBrush = new SolidBrush(Color.FromArgb(0, 173, 181));
            using var textBrush = Brushes.White;

            graphics.FillRectangle(accentBrush, 0, height - 6, width, 6);
            graphics.DrawString(fileName, titleFont, textBrush, new RectangleF(16, 16, width - 32, 80));
            if (!string.IsNullOrWhiteSpace(slideTitle))
            {
                graphics.DrawString(slideTitle, subtitleFont, textBrush, new RectangleF(16, 100, width - 32, height - 116));
            }
            else
            {
                graphics.DrawString("Anteprima non disponibile", subtitleFont, textBrush, new RectangleF(16, 100, width - 32, height - 116));
            }
        }

        bitmap.Save(cachePath, ImageFormat.Png);
        return PreviewResult.SuccessResult(cachePath);
    }
}
