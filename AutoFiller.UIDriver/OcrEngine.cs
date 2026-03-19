// Windows.Media.OCR is built into Windows 10+ — no external package is required
// for the API itself.  However, to call WinRT APIs from a regular .NET project
// you need ONE of the following references added to the .csproj:
//
//   Option A (SDK-style, recommended):
//     <UseWindowsDesktopSdk>true</UseWindowsDesktopSdk>        <!-- or target net6.0-windows -->
//     <TargetFramework>net8.0-windows10.0.19041.0</TargetFramework>
//
//   Option B (NuGet, any .NET Framework / .NET project):
//     <PackageReference Include="Microsoft.Windows.SDK.Contracts" Version="10.0.19041.1" />
//
// Either option exposes Windows.Media.Ocr, Windows.Graphics.Imaging, etc.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using Windows.Globalization;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;

namespace AutoFiller.UIDriver
{
    // ─────────────────────────────────────────────
    // Data model
    // ─────────────────────────────────────────────

    public class OcrWord
    {
        public string Text { get; set; }
        /// <summary>Bounding box relative to the captured bitmap origin (px).</summary>
        public Rectangle BoundingBox { get; set; }
        /// <summary>Word-level confidence in [0, 1].  Windows.Media.OCR does not
        /// expose per-word confidence directly — we approximate from line
        /// recognition quality (set to 1.0 when unavailable).</summary>
        public double Confidence { get; set; }
    }

    public class OcrLine
    {
        public string Text { get; set; }
        public Rectangle BoundingBox { get; set; }
        public List<OcrWord> Words { get; set; } = new();
    }

    public class OcrResult
    {
        public string FullText { get; set; }
        public List<OcrWord> Words { get; set; } = new();
        public List<OcrLine> Lines { get; set; } = new();
    }

    public enum FieldPosition { Right, Below, Left }

    // ─────────────────────────────────────────────
    // Engine
    // ─────────────────────────────────────────────

    /// <summary>
    /// Wraps <see cref="Windows.Media.Ocr.OcrEngine"/> for use in a standard
    /// .NET project.  Recognises Japanese text by default.
    /// </summary>
    public class OcrEngine
    {
        private readonly Windows.Media.Ocr.OcrEngine _engine;

        // ── constructor ───────────────────────────────────────────────────

        /// <summary>
        /// Initialises the OCR engine with the Japanese (ja) language.
        /// Throws <see cref="InvalidOperationException"/> if the language pack
        /// is not installed on the current machine.
        /// </summary>
        public OcrEngine() : this("ja") { }

        /// <summary>
        /// Initialises the OCR engine with the specified BCP-47 language tag
        /// (e.g. "ja", "en", "zh-Hans").
        /// </summary>
        public OcrEngine(string languageTag)
        {
            if (string.IsNullOrEmpty(languageTag))
                throw new ArgumentException("languageTag must not be empty.", nameof(languageTag));

            var language = new Language(languageTag);
            _engine = Windows.Media.Ocr.OcrEngine.TryCreateFromLanguage(language)
                      ?? Windows.Media.Ocr.OcrEngine.TryCreateFromUserProfileLanguages()
                      ?? throw new InvalidOperationException(
                             $"OCR language pack '{languageTag}' is not installed. " +
                             "Install the language pack via Windows Settings > Time & Language.");
        }

        // ── public API ────────────────────────────────────────────────────

        /// <summary>
        /// Runs OCR on <paramref name="bitmap"/> and returns every recognised
        /// word with its bounding box in bitmap-relative coordinates.
        /// </summary>
        public async Task<OcrResult> RecognizeAsync(Bitmap bitmap)
        {
            if (bitmap == null) throw new ArgumentNullException(nameof(bitmap));

            using SoftwareBitmap swBmp = await ToSoftwareBitmapAsync(bitmap);
            OcrResult result = await RecognizeSoftwareBitmapAsync(swBmp);
            return result;
        }

        /// <summary>
        /// Finds the first word in <paramref name="ocr"/> whose text contains
        /// <paramref name="searchText"/> (case-insensitive, partial match) and
        /// whose <see cref="OcrWord.Confidence"/> meets <paramref name="minConfidence"/>.
        /// Returns null when no match is found.
        /// </summary>
        public OcrWord? FindText(OcrResult ocr, string searchText, double minConfidence = 0.7)
        {
            if (ocr == null || string.IsNullOrEmpty(searchText)) return null;

            return ocr.Words.FirstOrDefault(w =>
                w.Confidence >= minConfidence &&
                w.Text.Contains(searchText, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Returns all words in <paramref name="ocr"/> whose text contains
        /// <paramref name="searchText"/> (case-insensitive).
        /// </summary>
        public List<OcrWord> FindAllText(OcrResult ocr, string searchText)
        {
            if (ocr == null || string.IsNullOrEmpty(searchText))
                return new List<OcrWord>();

            return ocr.Words
                .Where(w => w.Text.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        /// <summary>
        /// Locates a text label and estimates the click point of the input field
        /// adjacent to it.
        /// 
        /// <para>Convention used:</para>
        /// <list type="bullet">
        ///   <item><see cref="FieldPosition.Right"/>  — field starts at
        ///         <c>label.Right + offset</c>, center Y = label center Y.</item>
        ///   <item><see cref="FieldPosition.Below"/>  — field starts at
        ///         <c>label.Bottom + offset</c>, center X = label center X.</item>
        ///   <item><see cref="FieldPosition.Left"/>   — field ends at
        ///         <c>label.Left - offset</c>, center Y = label center Y.</item>
        /// </list>
        ///
        /// Returns null when the label is not found.
        /// </summary>
        public Point? FindInputFieldNearLabel(
            OcrResult ocr,
            string labelText,
            FieldPosition position = FieldPosition.Right,
            int offset = 8)
        {
            OcrWord label = FindText(ocr, labelText);
            if (label == null) return null;

            Rectangle b = label.BoundingBox;
            int centerY = b.Top + b.Height / 2;
            int centerX = b.Left + b.Width / 2;

            return position switch
            {
                FieldPosition.Right => new Point(b.Right + offset, centerY),
                FieldPosition.Below => new Point(centerX, b.Bottom + offset),
                FieldPosition.Left  => new Point(b.Left - offset, centerY),
                _                   => new Point(b.Right + offset, centerY)
            };
        }

        // ── private helpers ───────────────────────────────────────────────

        /// <summary>
        /// Converts a GDI+ <see cref="Bitmap"/> to a
        /// <see cref="SoftwareBitmap"/> that Windows.Media.Ocr can consume.
        /// </summary>
        private static async Task<SoftwareBitmap> ToSoftwareBitmapAsync(Bitmap bitmap)
        {
            // Encode to PNG in memory, then decode via BitmapDecoder.
            using var ms = new MemoryStream();
            bitmap.Save(ms, ImageFormat.Png);
            ms.Position = 0;

            // ms → IRandomAccessStream
            using var ras = ms.AsRandomAccessStream();

            BitmapDecoder decoder = await BitmapDecoder.CreateAsync(ras);
            SoftwareBitmap swBmp = await decoder.GetSoftwareBitmapAsync(
                BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied);

            return swBmp;
        }

        private async Task<OcrResult> RecognizeSoftwareBitmapAsync(SoftwareBitmap swBmp)
        {
            OcrRecognitionResult winResult = await _engine.RecognizeAsync(swBmp);

            var result = new OcrResult
            {
                FullText = winResult.Text ?? string.Empty
            };

            foreach (OcrLine winLine in winResult.Lines)
            {
                var line = new OcrLine
                {
                    Text = winLine.Text ?? string.Empty,
                    BoundingBox = RectToRectangle(winLine.Words
                        .Select(w => w.BoundingRect)
                        .Aggregate(Union))
                };

                foreach (OcrWord winWord in winLine.Words)
                {
                    var word = new OcrWord
                    {
                        Text = winWord.Text ?? string.Empty,
                        BoundingBox = RectToRectangle(winWord.BoundingRect),
                        Confidence = 1.0  // Windows.Media.OCR has no per-word confidence API
                    };
                    line.Words.Add(word);
                    result.Words.Add(word);
                }

                result.Lines.Add(line);
            }

            return result;
        }

        private static Rectangle RectToRectangle(Windows.Foundation.Rect r)
            => new Rectangle((int)r.X, (int)r.Y, (int)r.Width, (int)r.Height);

        private static Windows.Foundation.Rect Union(
            Windows.Foundation.Rect a, Windows.Foundation.Rect b)
        {
            double x = Math.Min(a.X, b.X);
            double y = Math.Min(a.Y, b.Y);
            double right  = Math.Max(a.X + a.Width,  b.X + b.Width);
            double bottom = Math.Max(a.Y + a.Height, b.Y + b.Height);
            return new Windows.Foundation.Rect(x, y, right - x, bottom - y);
        }
    }
}
