using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;   // for Clipboard (STA)

namespace AutoFiller.UIDriver
{
    // ─────────────────────────────────────────────
    // Grid layout model
    // ─────────────────────────────────────────────

    public class GridLayout
    {
        /// <summary>Absolute screen Y-coordinate of the column-header row.</summary>
        public int HeaderRowY { get; set; }
        /// <summary>Absolute screen Y-coordinate of the first data row (row 0).</summary>
        public int FirstDataRowY { get; set; }
        /// <summary>Estimated height of one data row in pixels.</summary>
        public int RowHeight { get; set; }
        public List<GridColumn> Columns { get; set; } = new List<GridColumn>();
    }

    public class GridColumn
    {
        public string HeaderText { get; set; }
        /// <summary>Absolute screen X-center of this column.</summary>
        public int CenterX { get; set; }
        public int Width { get; set; }
    }

    // ─────────────────────────────────────────────
    // VisionInteractor
    // ─────────────────────────────────────────────

    /// <summary>
    /// Vision-guided Win32 interactor.  Replaces UIAutomation entirely.
    /// Uses <see cref="ScreenCapture"/> + <see cref="OcrEngine"/> to locate
    /// controls, then drives input via mouse_event / keybd_event.
    /// </summary>
    public class VisionInteractor
    {
        // ── Win32 imports ─────────────────────────────────────────────────

        [DllImport("user32.dll")]
        private static extern void mouse_event(uint flags, int x, int y, int data, int extraInfo);

        [DllImport("user32.dll")]
        private static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hwnd);

        [DllImport("user32.dll")]
        private static extern void keybd_event(byte vk, byte scan, uint flags, int extraInfo);

        [DllImport("user32.dll")]
        private static extern short VkKeyScan(char ch);

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hwnd, out ScreenCapture.RECT rect);

        // mouse_event flags
        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP   = 0x0004;
        private const uint MOUSEEVENTF_WHEEL    = 0x0800;

        // keybd_event flags
        private const uint KEYEVENTF_KEYUP = 0x0002;

        // Virtual key codes
        private const byte VK_CONTROL = 0x11;
        private const byte VK_A       = 0x41;
        private const byte VK_V       = 0x56;
        private const byte VK_F2      = 0x71;
        private const byte VK_F9      = 0x78;
        private const byte VK_TAB     = 0x09;
        private const byte VK_RETURN  = 0x0D;

        // WM_KEYDOWN / WM_KEYUP
        private const uint WM_KEYDOWN = 0x0100;
        private const uint WM_KEYUP   = 0x0101;

        // Mouse wheel delta: one notch = 120 WHEEL_DELTA units
        private const int WHEEL_DELTA = 120;

        // How long to wait after a click before proceeding (ms).
        private const int DefaultClickDelayMs = 80;

        // ── fields ────────────────────────────────────────────────────────

        private readonly ScreenCapture _screen;
        private readonly OcrEngine _ocr;
        private IntPtr _hwnd = IntPtr.Zero;

        // Cached grid layout (invalidated on tab change).
        private GridLayout _cachedGrid;

        public VisionInteractor()
        {
            _screen = new ScreenCapture();
            _ocr    = new OcrEngine();          // Japanese OCR
        }

        public VisionInteractor(OcrEngine ocrEngine)
        {
            _screen = new ScreenCapture();
            _ocr    = ocrEngine ?? throw new ArgumentNullException(nameof(ocrEngine));
        }

        // ── public: attach ────────────────────────────────────────────────

        /// <summary>
        /// Finds the app window by partial title match and stores its handle.
        /// Returns false when no matching window is found.
        /// </summary>
        public bool Attach(string windowTitleContains)
        {
            if (string.IsNullOrEmpty(windowTitleContains))
                throw new ArgumentException("windowTitleContains must not be empty.",
                    nameof(windowTitleContains));

            _hwnd = _screen.FindWindowByTitle(windowTitleContains);
            _cachedGrid = null;
            return _hwnd != IntPtr.Zero;
        }

        // ── public: high-level ────────────────────────────────────────────

        /// <summary>
        /// Finds <paramref name="labelText"/> via OCR, clicks the adjacent
        /// input field, selects all existing content, then pastes
        /// <paramref name="value"/>.  Optionally verifies the fill with a
        /// second OCR pass.
        /// </summary>
        /// <returns>
        /// True when the value is visible in the area after filling;
        /// false when the label was not found or verification failed.
        /// </returns>
        public async Task<bool> FillFieldByLabel(
            string labelText,
            string value,
            FieldPosition fieldPosition = FieldPosition.Right)
        {
            EnsureAttached();

            // 1. Capture + OCR
            using Bitmap bmp = _screen.CaptureWindow(_hwnd);
            OcrResult ocr = await _ocr.RecognizeAsync(bmp);

            // 2. Locate the input field next to the label
            Point? fieldPoint = _ocr.FindInputFieldNearLabel(ocr, labelText, fieldPosition, offset: 8);
            if (!fieldPoint.HasValue) return false;

            // 3. Translate bitmap-relative coords to absolute screen coords
            Rectangle winBounds = _screen.GetWindowBounds(_hwnd);
            int screenX = winBounds.Left + fieldPoint.Value.X;
            int screenY = winBounds.Top  + fieldPoint.Value.Y;

            // 4. Click → SelectAll → Paste
            BringWindowToFront();
            await ClickAt(screenX, screenY);
            SelectAll();
            await PasteClipboard(value);

            // 5. Verify: re-OCR the label's bounding area + offset right/below
            OcrWord label = _ocr.FindText(ocr, labelText);
            if (label != null)
            {
                var region = BuildVerifyRegion(label.BoundingBox, fieldPosition, winBounds);
                return await VerifyValueInRegion(region, value);
            }

            // No label bounds → assume success (best-effort)
            return true;
        }

        /// <summary>
        /// Locates a tab by its header text, clicks it, and waits for the
        /// content pane to load.
        /// </summary>
        public async Task<bool> ClickTab(string tabText, int waitMs = 500)
        {
            EnsureAttached();

            using Bitmap bmp = _screen.CaptureWindow(_hwnd);
            OcrResult ocr = await _ocr.RecognizeAsync(bmp);

            OcrWord tab = _ocr.FindText(ocr, tabText);
            if (tab == null) return false;

            Rectangle winBounds = _screen.GetWindowBounds(_hwnd);
            int screenX = winBounds.Left + tab.BoundingBox.Left + tab.BoundingBox.Width  / 2;
            int screenY = winBounds.Top  + tab.BoundingBox.Top  + tab.BoundingBox.Height / 2;

            BringWindowToFront();
            await ClickAt(screenX, screenY);

            // Invalidate cached grid since the tab content changed.
            _cachedGrid = null;

            await Task.Delay(waitMs);
            return true;
        }

        /// <summary>
        /// Fills a single grid cell identified by its column header text and
        /// 0-based row index.
        /// </summary>
        public async Task<bool> FillGridCell(
            int rowIndex,
            string columnHeaderText,
            string value)
        {
            EnsureAttached();

            GridLayout layout = _cachedGrid ?? await DetectGrid();
            if (layout == null) return false;
            _cachedGrid = layout;

            GridColumn col = layout.Columns.FirstOrDefault(
                c => c.HeaderText != null &&
                     c.HeaderText.Contains(columnHeaderText, StringComparison.OrdinalIgnoreCase));
            if (col == null) return false;

            int cellScreenX = col.CenterX;
            int cellScreenY = layout.FirstDataRowY + rowIndex * layout.RowHeight
                              + layout.RowHeight / 2;

            BringWindowToFront();
            await ClickAt(cellScreenX, cellScreenY);
            SelectAll();
            await PasteClipboard(value);

            // Verify inside the cell area.
            var cellRegion = new Rectangle(
                col.CenterX - col.Width / 2,
                cellScreenY - layout.RowHeight / 2,
                col.Width,
                layout.RowHeight);

            return await VerifyValueInRegion(cellRegion, value);
        }

        /// <summary>
        /// Takes a screenshot, runs OCR, and infers the grid structure from
        /// a row of evenly-spaced column headers near the top visible area.
        /// </summary>
        public async Task<GridLayout?> DetectGrid()
        {
            EnsureAttached();

            using Bitmap bmp = _screen.CaptureWindow(_hwnd);
            OcrResult ocr = await _ocr.RecognizeAsync(bmp);
            Rectangle winBounds = _screen.GetWindowBounds(_hwnd);

            if (ocr.Words.Count == 0) return null;

            // Group words by approximate Y band (±8 px tolerance).
            var bands = GroupByY(ocr.Words, tolerance: 8);

            // The header row is the band with the highest density of evenly-spaced words.
            List<OcrWord> headerBand = FindHeaderBand(bands);
            if (headerBand == null || headerBand.Count < 2) return null;

            int headerBitmapY = headerBand.Average(w => w.BoundingBox.Top + w.BoundingBox.Height / 2.0)
                                          .RoundToInt();
            int headerHeight  = headerBand.Max(w => w.BoundingBox.Height);

            // Estimate row height: typically header height × 1.5, minimum 16 px.
            int rowHeight     = Math.Max(16, (int)(headerHeight * 1.5));
            int firstDataY    = headerBitmapY + headerHeight + 2;

            var columns = headerBand
                .OrderBy(w => w.BoundingBox.Left)
                .Select(w => new GridColumn
                {
                    HeaderText = w.Text,
                    CenterX    = winBounds.Left + w.BoundingBox.Left + w.BoundingBox.Width / 2,
                    Width      = w.BoundingBox.Width
                })
                .ToList();

            return new GridLayout
            {
                HeaderRowY    = winBounds.Top + headerBitmapY,
                FirstDataRowY = winBounds.Top + firstDataY,
                RowHeight     = rowHeight,
                Columns       = columns
            };
        }

        // ── public: verify ────────────────────────────────────────────────

        /// <summary>
        /// Captures <paramref name="screenRegion"/> (absolute screen coords),
        /// runs OCR, and checks whether <paramref name="expectedValue"/> appears.
        /// </summary>
        public async Task<bool> VerifyValueInRegion(Rectangle screenRegion, string expectedValue)
        {
            EnsureAttached();

            if (string.IsNullOrEmpty(expectedValue)) return true;

            Rectangle winBounds = _screen.GetWindowBounds(_hwnd);

            // Convert screen region to window-relative for CaptureRegion.
            var winRelRegion = new Rectangle(
                screenRegion.X - winBounds.Left,
                screenRegion.Y - winBounds.Top,
                screenRegion.Width,
                screenRegion.Height);

            // Guard: clamp to valid window area.
            if (winRelRegion.Width <= 0 || winRelRegion.Height <= 0) return false;

            try
            {
                using Bitmap regionBmp = _screen.CaptureRegion(_hwnd, winRelRegion);
                OcrResult regionOcr = await _ocr.RecognizeAsync(regionBmp);

                string normExpected = expectedValue.Trim();
                return regionOcr.Words.Any(w =>
                    w.Text.Contains(normExpected, StringComparison.OrdinalIgnoreCase));
            }
            catch
            {
                return false;   // region outside window or capture failed
            }
        }

        // ── public: low-level input ───────────────────────────────────────

        /// <summary>
        /// Sends a virtual key (e.g. F9, F2) to the attached window via
        /// <c>PostMessage</c> for reliability with background windows.
        /// </summary>
        public void SendFunctionKey(byte vkKey)
        {
            EnsureAttached();
            PostMessage(_hwnd, WM_KEYDOWN, (IntPtr)vkKey, IntPtr.Zero);
            Thread.Sleep(30);
            PostMessage(_hwnd, WM_KEYUP,   (IntPtr)vkKey, IntPtr.Zero);
        }

        /// <summary>
        /// Scrolls the app by sending a mouse-wheel event over the current
        /// cursor position.  Positive <paramref name="deltaRows"/> scrolls up;
        /// negative scrolls down.
        /// </summary>
        public async Task ScrollGrid(int deltaRows)
        {
            EnsureAttached();
            BringWindowToFront();
            await Task.Delay(50);

            int wheelData = deltaRows * WHEEL_DELTA;
            mouse_event(MOUSEEVENTF_WHEEL, 0, 0, wheelData, 0);
            await Task.Delay(80);
        }

        // ── private: low-level input ──────────────────────────────────────

        /// <summary>
        /// Moves the cursor to (<paramref name="screenX"/>, <paramref name="screenY"/>)
        /// and performs a single left click.
        /// </summary>
        private async Task ClickAt(int screenX, int screenY)
        {
            SetCursorPos(screenX, screenY);
            await Task.Delay(30);
            mouse_event(MOUSEEVENTF_LEFTDOWN, screenX, screenY, 0, 0);
            await Task.Delay(DefaultClickDelayMs);
            mouse_event(MOUSEEVENTF_LEFTUP,   screenX, screenY, 0, 0);
            await Task.Delay(DefaultClickDelayMs);
        }

        /// <summary>Sends Ctrl+A to select all text in the focused control.</summary>
        private void SelectAll()
        {
            keybd_event(VK_CONTROL, 0, 0,              0);
            keybd_event(VK_A,       0, 0,              0);
            keybd_event(VK_A,       0, KEYEVENTF_KEYUP, 0);
            keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);
            Thread.Sleep(30);
        }

        /// <summary>
        /// Sets <paramref name="text"/> on the system clipboard (requires STA)
        /// and sends Ctrl+V.  Automatically marshals to an STA thread when called
        /// from an MTA context.
        /// </summary>
        private async Task PasteClipboard(string text)
        {
            // Clipboard API requires STA thread.
            if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
            {
                Clipboard.SetText(text);
            }
            else
            {
                var tcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
                var sta = new Thread(() =>
                {
                    try
                    {
                        Clipboard.SetText(text);
                        tcs.SetResult(true);
                    }
                    catch (Exception ex)
                    {
                        tcs.SetException(ex);
                    }
                });
                sta.SetApartmentState(ApartmentState.STA);
                sta.Start();
                await tcs.Task;
            }

            await Task.Delay(30);   // give clipboard time to settle

            keybd_event(VK_CONTROL, 0, 0,               0);
            keybd_event(VK_V,       0, 0,               0);
            keybd_event(VK_V,       0, KEYEVENTF_KEYUP,  0);
            keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP,  0);
            await Task.Delay(60);
        }

        // ── private: helpers ──────────────────────────────────────────────

        private void BringWindowToFront()
        {
            if (_hwnd != IntPtr.Zero)
                SetForegroundWindow(_hwnd);
            Thread.Sleep(80);
        }

        private void EnsureAttached()
        {
            if (_hwnd == IntPtr.Zero)
                throw new InvalidOperationException(
                    "VisionInteractor is not attached to a window. Call Attach() first.");
        }

        /// <summary>
        /// Groups OCR words into horizontal bands where all words share roughly
        /// the same Y centre (within <paramref name="tolerance"/> pixels).
        /// </summary>
        private static List<List<OcrWord>> GroupByY(
            IEnumerable<OcrWord> words, int tolerance)
        {
            var bands = new List<List<OcrWord>>();
            foreach (var word in words.OrderBy(w => w.BoundingBox.Top))
            {
                int wordCy = word.BoundingBox.Top + word.BoundingBox.Height / 2;
                var band = bands.FirstOrDefault(b =>
                {
                    int bandCy = b[0].BoundingBox.Top + b[0].BoundingBox.Height / 2;
                    return Math.Abs(wordCy - bandCy) <= tolerance;
                });

                if (band != null)
                    band.Add(word);
                else
                    bands.Add(new List<OcrWord> { word });
            }
            return bands;
        }

        /// <summary>
        /// Identifies the band most likely to be a column-header row:
        /// highest word count with even horizontal spacing.
        /// </summary>
        private static List<OcrWord> FindHeaderBand(List<List<OcrWord>> bands)
        {
            List<OcrWord> best = null;
            double bestScore = -1;

            foreach (var band in bands)
            {
                if (band.Count < 2) continue;

                var sorted = band.OrderBy(w => w.BoundingBox.Left).ToList();
                var gaps = new List<double>();
                for (int i = 1; i < sorted.Count; i++)
                    gaps.Add(sorted[i].BoundingBox.Left - sorted[i - 1].BoundingBox.Right);

                double avgGap      = gaps.Average();
                double gapVariance = gaps.Average(g => Math.Abs(g - avgGap));

                // Score: many words + consistent gaps = likely a header row.
                double score = band.Count * 2.0 - gapVariance / Math.Max(1, avgGap);

                if (score > bestScore)
                {
                    bestScore = score;
                    best = band;
                }
            }

            return best;
        }

        /// <summary>
        /// Builds a verification region in absolute screen coordinates based on
        /// the label's bounding box and the expected field position.
        /// </summary>
        private static Rectangle BuildVerifyRegion(
            Rectangle labelBitmapBox,
            FieldPosition position,
            Rectangle winBounds)
        {
            const int verifyW = 200;
            const int verifyH = 32;

            int labelScreenLeft = winBounds.Left + labelBitmapBox.Left;
            int labelScreenTop  = winBounds.Top  + labelBitmapBox.Top;

            return position switch
            {
                FieldPosition.Below =>
                    new Rectangle(labelScreenLeft, labelScreenTop + labelBitmapBox.Height + 2, verifyW, verifyH),
                FieldPosition.Left =>
                    new Rectangle(labelScreenLeft - verifyW - 4, labelScreenTop, verifyW, verifyH),
                _ =>
                    new Rectangle(labelScreenLeft + labelBitmapBox.Width + 4, labelScreenTop, verifyW, verifyH)
            };
        }
    }

    // ── extension helpers ────────────────────────────────────────────────────

    internal static class DoubleExtensions
    {
        internal static int RoundToInt(this double value) => (int)Math.Round(value);
    }
}
