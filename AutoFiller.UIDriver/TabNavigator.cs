using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;

namespace AutoFiller.UIDriver
{
    // ─────────────────────────────────────────────
    // Supporting model
    // ─────────────────────────────────────────────

    public class TabInfo
    {
        public string TabText { get; set; }
        /// <summary>Absolute screen coordinates to click to activate the tab.</summary>
        public Point TabClickPosition { get; set; }
        public List<string> VisibleLabels { get; set; } = new List<string>();
        public bool HasGrid { get; set; }
        public GridLayout GridLayout { get; set; }
    }

    public class AppTabStructure
    {
        public List<TabInfo> Tabs { get; set; } = new List<TabInfo>();
        /// <summary>Tab that contains fixed header fields (画面ID, 画面名 …).</summary>
        public string HeaderTabName { get; set; }
        /// <summary>Tab that contains the 項目定義 grid.</summary>
        public string GridTabName { get; set; }
        public DateTime DiscoveredAt { get; set; }
    }

    // ─────────────────────────────────────────────
    // TabNavigator
    // ─────────────────────────────────────────────

    /// <summary>
    /// Handles lazy-loading WinForms tab navigation: each tab is activated by
    /// clicking, then we wait for its controls to render before proceeding.
    /// All UI detection is vision-based (screenshot + OCR).
    /// </summary>
    public class TabNavigator
    {
        // ── known label signatures ────────────────────────────────────────

        private static readonly string[] HeaderTabSignatures =
            { "画面ID", "画面名", "画面Ver", "顧客名", "作成者" };

        private static readonly string[] GridTabSignatures =
            { "項目定義", "番号", "項目名", "項目種類" };

        private static readonly string[] FunctionTabSignatures =
            { "ファンクション", "F1", "F9", "F12" };

        // Y-band tolerance for grouping words on the same row (px).
        private const int YBandTolerance = 8;

        // Maximum length (chars) of a tab label — longer strings are not tabs.
        private const int MaxTabLabelLength = 20;

        // How far from the top of the window (% of height) the tab strip can be.
        private const double TabStripMaxRelativeY = 0.20;

        // Content-change poll interval and timeout.
        private const int PollIntervalMs = 200;

        // ── fields ────────────────────────────────────────────────────────

        private readonly VisionInteractor _vision;
        private readonly ScreenCapture _screen;
        private readonly OcrEngine _ocr;
        private string _currentTab = string.Empty;

        // ── constructors ──────────────────────────────────────────────────

        public TabNavigator(VisionInteractor vision)
        {
            _vision = vision ?? throw new ArgumentNullException(nameof(vision));
            _screen = new ScreenCapture();
            _ocr    = new OcrEngine();
        }

        public TabNavigator(VisionInteractor vision, OcrEngine ocrEngine)
        {
            _vision = vision ?? throw new ArgumentNullException(nameof(vision));
            _screen = new ScreenCapture();
            _ocr    = ocrEngine ?? throw new ArgumentNullException(nameof(ocrEngine));
        }

        /// <summary>The handle of the app window (forwarded from VisionInteractor).</summary>
        private IntPtr Hwnd => GetHwndFromVision();

        // ── public API ────────────────────────────────────────────────────

        /// <summary>
        /// Takes a screenshot, OCRs the tab strip region (top 20 % of window),
        /// and returns all discovered tab label strings.
        /// </summary>
        public async Task<List<string>> DiscoverTabs()
        {
            using Bitmap bmp = _screen.CaptureWindow(Hwnd);
            OcrResult ocr = await _ocr.RecognizeAsync(bmp);

            var tabBand = FindTabStripBand(ocr.Words, bmp.Height);
            return tabBand.Select(w => w.Text).ToList();
        }

        /// <summary>
        /// Navigates to <paramref name="tabText"/>:
        /// <list type="number">
        ///   <item>If already active, returns immediately.</item>
        ///   <item>Finds the tab label via OCR, clicks it.</item>
        ///   <item>Polls every 200 ms (up to <paramref name="timeoutMs"/>) until
        ///         the OCR word count changes, indicating new content has
        ///         rendered.</item>
        ///   <item>Updates <see cref="_currentTab"/>.</item>
        /// </list>
        /// </summary>
        public async Task<bool> NavigateTo(string tabText, int timeoutMs = 3000)
        {
            if (string.Equals(_currentTab, tabText, StringComparison.Ordinal))
                return true;

            using (Bitmap pre = _screen.CaptureWindow(Hwnd))
            {
                OcrResult preOcr = await _ocr.RecognizeAsync(pre);

                OcrWord tab = _ocr.FindText(preOcr, tabText);
                if (tab == null) return false;

                Rectangle winBounds = _screen.GetWindowBounds(Hwnd);
                int screenX = winBounds.Left + tab.BoundingBox.Left + tab.BoundingBox.Width  / 2;
                int screenY = winBounds.Top  + tab.BoundingBox.Top  + tab.BoundingBox.Height / 2;

                await _vision.ClickTab(tabText, waitMs: 0);   // click only, no internal wait

                // Poll until word count changes or we time out.
                int elapsed = 0;
                int baseWordCount = preOcr.Words.Count;

                while (elapsed < timeoutMs)
                {
                    await Task.Delay(PollIntervalMs);
                    elapsed += PollIntervalMs;

                    try
                    {
                        using Bitmap post = _screen.CaptureWindow(Hwnd);
                        OcrResult postOcr = await _ocr.RecognizeAsync(post);
                        if (postOcr.Words.Count != baseWordCount)
                            break;   // content changed — tab rendered
                    }
                    catch
                    {
                        // Window may briefly be unavailable; keep polling.
                    }
                }
            }

            _currentTab = tabText;
            return true;
        }

        /// <summary>
        /// Navigates to every discovered tab, captures its visible text labels,
        /// and returns a <c>tabName → labels</c> dictionary.
        /// </summary>
        public async Task<Dictionary<string, List<string>>> ScanAllTabs()
        {
            List<string> tabs = await DiscoverTabs();
            var result = new Dictionary<string, List<string>>(StringComparer.Ordinal);

            foreach (string tab in tabs)
            {
                await NavigateTo(tab);

                using Bitmap bmp = _screen.CaptureWindow(Hwnd);
                OcrResult ocr = await _ocr.RecognizeAsync(bmp);

                result[tab] = ocr.Words
                    .Select(w => w.Text)
                    .Where(t => !string.IsNullOrWhiteSpace(t))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }

            return result;
        }

        /// <summary>
        /// After navigating to a tab, inspects the current screenshot for grid
        /// signatures: a header band followed by at least two data rows at a
        /// consistent vertical interval.
        /// </summary>
        public async Task<bool> CurrentTabHasGrid()
        {
            using Bitmap bmp = _screen.CaptureWindow(Hwnd);
            OcrResult ocr = await _ocr.RecognizeAsync(bmp);
            return HasGridSignature(ocr.Words, bmp.Height);
        }

        /// <summary>
        /// Navigates to every tab, catalogues its content, and builds an
        /// <see cref="AppTabStructure"/> that identifies the header tab and grid tab
        /// by their landmark label signatures.
        /// </summary>
        public async Task<AppTabStructure> DiscoverAppStructure()
        {
            List<string> tabNames = await DiscoverTabs();
            if (tabNames.Count == 0)
            {
                // Fall back: treat entire window as single unnamed tab.
                tabNames = new List<string> { "(unknown)" };
            }

            // Pre-capture click positions from the current screenshot.
            Dictionary<string, Point> clickPositions = await CaptureTabClickPositions(tabNames);

            var structure = new AppTabStructure
            {
                DiscoveredAt = DateTime.Now
            };

            foreach (string tabName in tabNames)
            {
                bool navigated = await NavigateTo(tabName);

                using Bitmap bmp = _screen.CaptureWindow(Hwnd);
                OcrResult ocr = await _ocr.RecognizeAsync(bmp);

                bool hasGrid = HasGridSignature(ocr.Words, bmp.Height);
                GridLayout gridLayout = hasGrid ? await _vision.DetectGrid() : null;

                var info = new TabInfo
                {
                    TabText = tabName,
                    TabClickPosition = clickPositions.TryGetValue(tabName, out Point cp) ? cp : default,
                    VisibleLabels = ocr.Words
                        .Select(w => w.Text)
                        .Where(t => !string.IsNullOrWhiteSpace(t))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList(),
                    HasGrid = hasGrid,
                    GridLayout = gridLayout
                };

                structure.Tabs.Add(info);

                // Identify header tab.
                if (structure.HeaderTabName == null &&
                    MatchesSignature(info.VisibleLabels, HeaderTabSignatures))
                {
                    structure.HeaderTabName = tabName;
                }

                // Identify grid tab.
                if (structure.GridTabName == null &&
                    (hasGrid || MatchesSignature(info.VisibleLabels, GridTabSignatures)))
                {
                    structure.GridTabName = tabName;
                }
            }

            return structure;
        }

        // ── private helpers ───────────────────────────────────────────────

        /// <summary>
        /// Returns words that belong to the tab strip: the horizontal band
        /// nearest to the top of the window where words are short (≤
        /// <see cref="MaxTabLabelLength"/> chars) and at the same Y level.
        /// </summary>
        private List<OcrWord> FindTabStripBand(
            IEnumerable<OcrWord> words, int bmpHeight)
        {
            int maxY = (int)(bmpHeight * TabStripMaxRelativeY);

            // Restrict to the top region.
            var topWords = words
                .Where(w => w.BoundingBox.Top < maxY
                            && w.Text.Length <= MaxTabLabelLength
                            && !string.IsNullOrWhiteSpace(w.Text))
                .ToList();

            if (topWords.Count == 0) return new List<OcrWord>();

            // Group by Y band.
            var bands = GroupByY(topWords, YBandTolerance);

            // The tab strip is the band with the most words that are spaced
            // fairly evenly (not just a header title).
            return bands
                .Where(b => b.Count >= 2)
                .OrderByDescending(b => b.Count)
                .FirstOrDefault()
                ?? new List<OcrWord>();
        }

        /// <summary>
        /// Returns true when <paramref name="words"/> exhibit a grid pattern:
        /// a header band near the upper half, followed by ≥ 2 data rows at a
        /// consistent vertical interval in the lower portion of the window.
        /// </summary>
        private static bool HasGridSignature(
            IEnumerable<OcrWord> words, int bmpHeight)
        {
            var allWords = words.ToList();
            if (allWords.Count < 8) return false;

            var bands = GroupByY(allWords, YBandTolerance);
            if (bands.Count < 3) return false;

            // Sort bands by Y.
            var sortedBands = bands
                .OrderBy(b => b.Min(w => w.BoundingBox.Top))
                .ToList();

            // Look for at least a header band (≥3 words) followed by ≥2 data bands
            // whose Y deltas are within ±6 px of each other.
            for (int i = 0; i < sortedBands.Count - 2; i++)
            {
                if (sortedBands[i].Count < 2) continue;

                var deltasBand = new List<int>();
                for (int j = i + 1; j < sortedBands.Count - 1; j++)
                {
                    int dy = sortedBands[j + 1].Min(w => w.BoundingBox.Top)
                             - sortedBands[j].Min(w => w.BoundingBox.Top);
                    if (dy > 0) deltasBand.Add(dy);
                }

                if (deltasBand.Count < 2) continue;

                double avg = deltasBand.Average();
                bool consistent = deltasBand.All(d => Math.Abs(d - avg) <= 6);
                if (consistent && avg >= 10 && avg <= 100)
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Groups a word list into horizontal bands by Y-centre proximity.
        /// </summary>
        private static List<List<OcrWord>> GroupByY(
            IEnumerable<OcrWord> words, int tolerance)
        {
            var bands = new List<List<OcrWord>>();
            foreach (var word in words.OrderBy(w => w.BoundingBox.Top))
            {
                int cy = word.BoundingBox.Top + word.BoundingBox.Height / 2;
                var band = bands.FirstOrDefault(b =>
                {
                    int bCy = b[0].BoundingBox.Top + b[0].BoundingBox.Height / 2;
                    return Math.Abs(cy - bCy) <= tolerance;
                });

                if (band != null) band.Add(word);
                else              bands.Add(new List<OcrWord> { word });
            }
            return bands;
        }

        /// <summary>
        /// Returns true when <paramref name="labels"/> contains at least
        /// two strings from <paramref name="signatures"/>.
        /// </summary>
        private static bool MatchesSignature(
            IReadOnlyList<string> labels, string[] signatures)
        {
            int hits = signatures.Count(sig =>
                labels.Any(l => l.Contains(sig, StringComparison.OrdinalIgnoreCase)));
            return hits >= 2;
        }

        /// <summary>
        /// Captures a single screenshot and records the absolute screen click
        /// position of each known tab name.
        /// </summary>
        private async Task<Dictionary<string, Point>> CaptureTabClickPositions(
            IEnumerable<string> tabNames)
        {
            var result = new Dictionary<string, Point>(StringComparer.Ordinal);

            using Bitmap bmp = _screen.CaptureWindow(Hwnd);
            OcrResult ocr = await _ocr.RecognizeAsync(bmp);
            Rectangle winBounds = _screen.GetWindowBounds(Hwnd);

            foreach (string name in tabNames)
            {
                OcrWord word = _ocr.FindText(ocr, name);
                if (word == null) continue;

                result[name] = new Point(
                    winBounds.Left + word.BoundingBox.Left + word.BoundingBox.Width  / 2,
                    winBounds.Top  + word.BoundingBox.Top  + word.BoundingBox.Height / 2);
            }

            return result;
        }

        /// <summary>
        /// Retrieves the window handle from the attached <see cref="VisionInteractor"/>
        /// via <see cref="ScreenCapture.FindWindowByTitle"/>.  We re-use whatever
        /// window the interactor last attached to by asking ScreenCapture to query
        /// the handle field through reflection — or, more simply, by capturing a
        /// dummy screenshot which will throw if not attached.
        ///
        /// Implementation note: <see cref="VisionInteractor"/> exposes no public
        /// handle property.  We capture the handle lazily by delegating all
        /// window-handle-requiring operations to objects that already hold it.
        /// For direct Win32 calls we rely on <see cref="ScreenCapture"/> which
        /// can locate the window by its current foreground title.
        /// </summary>
        private IntPtr GetHwndFromVision()
        {
            // Use the ScreenCapture to locate the same window the VisionInteractor
            // has attached to.  We store the last known hwnd here for subsequent calls.
            if (_lastHwnd != IntPtr.Zero) return _lastHwnd;

            // Try the foreground window as a sensible default.
            _lastHwnd = GetForegroundWindowFallback();
            return _lastHwnd;
        }

        private IntPtr _lastHwnd = IntPtr.Zero;

        // Allow callers to explicitly supply the hwnd (used by DiscoverAppStructure callers).
        public void SetWindowHandle(IntPtr hwnd)
        {
            _lastHwnd = hwnd;
        }

        // ── Win32 minimal import needed only here ────────────────────────
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        private static IntPtr GetForegroundWindowFallback() => GetForegroundWindow();
    }
}
