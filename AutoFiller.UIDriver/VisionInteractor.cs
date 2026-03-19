using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;   // for Clipboard (STA)
using Microsoft.Extensions.Logging;

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

        [StructLayout(LayoutKind.Sequential)]
        private struct MOUSEINPUT
        {
            public int    dx, dy, mouseData;
            public uint   dwFlags, time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct INPUT
        {
            public uint      type;
            public MOUSEINPUT mi;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        [DllImport("user32.dll", EntryPoint = "GetSystemMetrics")]
        private static extern int GetSysMetrics(int nIndex);

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

        // INPUT type
        private const uint INPUT_MOUSE = 0;

        // MOUSEEVENTF flags
        private const uint MOUSEEVENTF_MOVE      = 0x0001;
        private const uint MOUSEEVENTF_LEFTDOWN  = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP    = 0x0004;
        private const uint MOUSEEVENTF_WHEEL     = 0x0800;
        private const uint MOUSEEVENTF_ABSOLUTE  = 0x8000;

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

        // ── fields ────────────────────────────────────────────────────────

        private readonly ScreenCapture _screen;
        private readonly OcrEngine _ocr;
        private readonly TimingConfig _timing;
        private readonly ILogger<VisionInteractor> _logger;
        private IntPtr _hwnd = IntPtr.Zero;

        // Cached grid layout (invalidated on tab change).
        private GridLayout _cachedGrid;

        // Column header → cell type ("text" | "dropdown"). Populated via SetColumnCellTypes().
        private IReadOnlyDictionary<string, string> _columnCellTypes;

        public VisionInteractor()
            : this(OcrEngineProvider.Instance) { }

        public VisionInteractor(
            OcrEngine ocr,
            TimingConfig timing = null,
            ILogger<VisionInteractor> logger = null)
        {
            _screen = new ScreenCapture();
            _ocr    = ocr ?? OcrEngineProvider.Instance;
            _timing = timing ?? TimingConfig.Default;
            _logger = logger ?? NullLogger<VisionInteractor>.Instance;
        }

        // ── public: attach ── + mapping hints ─────────────────────────────

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

        /// <summary>
        /// Provides per-column cell-type hints used by <see cref="FillGridCell"/>
        /// to choose between clipboard paste and dropdown selection.
        /// Keys are column header strings (case-insensitive match).
        /// Values should be <c>"text"</c> or <c>"dropdown"</c>.
        /// </summary>
        public void SetColumnCellTypes(IReadOnlyDictionary<string, string> types)
        {
            _columnCellTypes = types;
        }

        // ── public: high-level ────────────────────────────────────────────

        /// <summary>
        /// Finds <paramref name="labelText"/> via OCR, clicks the adjacent
        /// input field, selects all existing content, then pastes
        /// <paramref name="value"/>. Retries up to 3 times on failure.
        /// </summary>
        /// <returns>
        /// True when the value is visible in the area after filling;
        /// false when the label was not found after all retry attempts.
        /// </returns>
        public async Task<bool> FillFieldByLabel(
            string labelText,
            string value,
            FieldPosition fieldPosition = FieldPosition.Right)
            => await RetryAsync(() => FillFieldByLabelOnce(labelText, value, fieldPosition));

        private async Task<bool> FillFieldByLabelOnce(
            string labelText,
            string value,
            FieldPosition fieldPosition)
        {
            EnsureAttached();

            _logger.LogDebug("Searching for label '{Label}'", labelText);

            // 1. Capture + OCR
            using Bitmap bmp = _screen.CaptureWindow(_hwnd);
            OcrResult ocr = await _ocr.RecognizeAsync(bmp);

            // 2. Locate the input field next to the label
            Point? fieldPoint = _ocr.FindInputFieldNearLabel(ocr, labelText, fieldPosition, offset: 8);
            if (!fieldPoint.HasValue)
            {
                _logger.LogWarning("Label '{Label}' not found in screenshot", labelText);
                return false;
            }

            // 3. Translate bitmap-relative coords to absolute screen coords
            Rectangle winBounds = _screen.GetWindowBounds(_hwnd);
            int screenX = winBounds.Left + fieldPoint.Value.X;
            int screenY = winBounds.Top  + fieldPoint.Value.Y;

            _logger.LogDebug("Label '{Label}' found, clicking field at ({FX},{FY})",
                labelText, screenX, screenY);

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
        /// 0-based row index. Retries up to 3 times on failure.
        /// Uses dropdown selection for columns flagged as <c>"dropdown"</c>
        /// via <see cref="SetColumnCellTypes"/>.
        /// </summary>
        public async Task<bool> FillGridCell(
            int rowIndex,
            string columnHeaderText,
            string value)
            => await RetryAsync(() => FillGridCellOnce(rowIndex, columnHeaderText, value));

        private async Task<bool> FillGridCellOnce(
            int rowIndex,
            string columnHeaderText,
            string value)
        {
            EnsureAttached();

            _logger.LogDebug("Filling grid row={Row} col='{Col}' value='{Val}'",
                rowIndex, columnHeaderText, value);

            GridLayout layout = _cachedGrid ?? await DetectGrid();
            if (layout == null) return false;
            _cachedGrid = layout;

            GridColumn col = layout.Columns.FirstOrDefault(
                c => c.HeaderText != null &&
                     c.HeaderText.Contains(columnHeaderText, StringComparison.OrdinalIgnoreCase));
            if (col == null)
            {
                _logger.LogWarning("Grid column '{Col}' not found in layout", columnHeaderText);
                return false;
            }

            int cellScreenX = col.CenterX;
            int cellScreenY = layout.FirstDataRowY + rowIndex * layout.RowHeight
                              + layout.RowHeight / 2;

            // Use dropdown selection when the column is flagged as "dropdown".
            if (GetCellType(columnHeaderText) == "dropdown")
                return await SelectDropdownValue(cellScreenX, cellScreenY, value);

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
            var bands = OcrUtils.GroupByY(ocr.Words, 8);

            // The header row is the band with the highest density of evenly-spaced words.
            List<OcrWord> headerBand = OcrUtils.FindHeaderBand(bands);
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

        // ── public: dropdown ──────────────────────────────────────────────

        /// <summary>
        /// Clicks (<paramref name="screenX"/>, <paramref name="screenY"/>) to open
        /// a dropdown/combobox, then polls the region below the click for up to
        /// <paramref name="timeoutMs"/> ms until OCR finds <paramref name="targetValue"/>
        /// in the popup list, and clicks the matching option.
        /// </summary>
        /// <returns>True when the option was found and clicked; false on timeout.</returns>
        public async Task<bool> SelectDropdownValue(
            int screenX, int screenY,
            string targetValue,
            int timeoutMs = 2000)
        {
            EnsureAttached();

            _logger.LogDebug("Opening dropdown at ({X},{Y}), looking for '{Val}'",
                screenX, screenY, targetValue);

            await ClickAt(screenX, screenY);
            await Task.Delay(200);

            // Search in a region below the click point for the popup list.
            var searchRegion = new Rectangle(screenX - 100, screenY, 200, 300);
            var sw = Stopwatch.StartNew();

            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                Rectangle winBounds = _screen.GetWindowBounds(_hwnd);
                var winRelRegion = new Rectangle(
                    searchRegion.X - winBounds.Left,
                    searchRegion.Y - winBounds.Top,
                    searchRegion.Width,
                    searchRegion.Height);
                winRelRegion.Intersect(new Rectangle(0, 0, winBounds.Width, winBounds.Height));

                if (winRelRegion.Width > 0 && winRelRegion.Height > 0)
                {
                    using var regionBmp = _screen.CaptureRegion(_hwnd, winRelRegion);
                    OcrResult popupOcr = await _ocr.RecognizeAsync(regionBmp);

                    OcrWord match = _ocr.FindText(popupOcr, targetValue);
                    if (match != null)
                    {
                        int absX = searchRegion.X + match.BoundingBox.Left + match.BoundingBox.Width  / 2;
                        int absY = searchRegion.Y + match.BoundingBox.Top  + match.BoundingBox.Height / 2;
                        await ClickAt(absX, absY);
                        return true;
                    }
                }

                await Task.Delay(200);
            }

            _logger.LogWarning("Dropdown option '{Val}' not found within {Timeout}ms",
                targetValue, timeoutMs);
            return false;
        }

        /// <summary>
        /// Re-detects grid column positions from the live app screenshot and
        /// updates <see cref="GridColumnMapping.GridColumnX"/> for every column
        /// in <paramref name="config"/>. Call this after loading a
        /// <see cref="MappingConfig"/> from JSON (where <c>GridColumnX</c> was
        /// not persisted) or when the app window may have moved.
        /// </summary>
        public async Task RefreshGridColumnPositions(AutoFiller.Core.MappingConfig config)
        {
            if (config?.Grid?.Columns == null) return;

            GridLayout layout = await DetectGrid();
            if (layout == null)
            {
                _logger.LogWarning("RefreshGridColumnPositions: DetectGrid returned null");
                return;
            }

            foreach (var kv in config.Grid.Columns)
            {
                GridColumn gridCol = layout.Columns.FirstOrDefault(c =>
                    c.HeaderText != null &&
                    c.HeaderText.Contains(kv.Key, StringComparison.OrdinalIgnoreCase));

                if (gridCol != null)
                {
                    kv.Value.GridColumnX = gridCol.CenterX;
                    _logger.LogDebug("Refreshed column '{Col}' X → {X}", kv.Key, gridCol.CenterX);
                }
                else
                {
                    _logger.LogWarning("RefreshGridColumnPositions: column '{Col}' not found in OCR layout", kv.Key);
                }
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
            Thread.Sleep(_timing.AfterFunctionKeyMs);
            PostMessage(_hwnd, WM_KEYUP,   (IntPtr)vkKey, IntPtr.Zero);
        }

        /// <summary>
        /// Scrolls the grid. Positive <paramref name="deltaRows"/> scrolls DOWN
        /// (shows lower rows). Negative scrolls UP (shows higher rows).
        /// One unit = one row height ≈ one WHEEL_DELTA notch.
        /// </summary>
        public async Task ScrollGrid(int deltaRows)
        {
            EnsureAttached();
            BringWindowToFront();
            await Task.Delay(_timing.ClickSettleMs);

            // Positive deltaRows = scroll down = NEGATIVE wheel delta.
            int wheelData = -deltaRows * WHEEL_DELTA;
            var input = new INPUT
            {
                type = INPUT_MOUSE,
                mi   = new MOUSEINPUT { dwFlags = MOUSEEVENTF_WHEEL, mouseData = wheelData }
            };
            SendInput(1, new[] { input }, Marshal.SizeOf(typeof(INPUT)));
            await Task.Delay(_timing.AfterScrollMs);
        }

        // ── private: low-level input ──────────────────────────────────────

        /// <summary>
        /// Moves the cursor to (<paramref name="screenX"/>, <paramref name="screenY"/>)
        /// and performs a single left click.
        /// </summary>
        private async Task ClickAt(int screenX, int screenY)
        {
            int screenW = GetSysMetrics(0);   // SM_CXSCREEN
            int screenH = GetSysMetrics(1);   // SM_CYSCREEN
            int normX   = screenX * 65535 / screenW;
            int normY   = screenY * 65535 / screenH;

            var inputs = new INPUT[]
            {
                new INPUT { type = INPUT_MOUSE, mi = new MOUSEINPUT {
                    dx = normX, dy = normY,
                    dwFlags = MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE } },
                new INPUT { type = INPUT_MOUSE, mi = new MOUSEINPUT {
                    dwFlags = MOUSEEVENTF_LEFTDOWN } },
                new INPUT { type = INPUT_MOUSE, mi = new MOUSEINPUT {
                    dwFlags = MOUSEEVENTF_LEFTUP } }
            };

            BringWindowToFront();
            await Task.Delay(_timing.ClickSettleMs);
            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
            await Task.Delay(_timing.ClickDelayMs);
        }

        /// <summary>Sends Ctrl+A to select all text in the focused control.</summary>
        private void SelectAll()
        {
            keybd_event(VK_CONTROL, 0, 0,              0);
            keybd_event(VK_A,       0, 0,              0);
            keybd_event(VK_A,       0, KEYEVENTF_KEYUP, 0);
            keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);
            Thread.Sleep(_timing.SelectAllMs);
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

            await Task.Delay(_timing.ClipboardSettleMs);   // give clipboard time to settle

            keybd_event(VK_CONTROL, 0, 0,               0);
            keybd_event(VK_V,       0, 0,               0);
            keybd_event(VK_V,       0, KEYEVENTF_KEYUP,  0);
            keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP,  0);
            await Task.Delay(_timing.AfterPasteMs);
        }

        // ── private: helpers ──────────────────────────────────────────────

        private void BringWindowToFront()
        {
            if (_hwnd != IntPtr.Zero)
                SetForegroundWindow(_hwnd);
            Thread.Sleep(_timing.BringToFrontMs);
        }

        private void EnsureAttached()
        {
            if (_hwnd == IntPtr.Zero)
                throw new InvalidOperationException(
                    "VisionInteractor is not attached to a window. Call Attach() first.");
        }

        /// <summary>
        /// Retries <paramref name="action"/> up to <paramref name="maxAttempts"/> times,
        /// waiting <paramref name="retryDelayMs"/> ms between attempts.
        /// Returns true on first success, false after all attempts exhausted.
        /// </summary>
        private static async Task<bool> RetryAsync(
            Func<Task<bool>> action,
            int maxAttempts  = 3,
            int retryDelayMs = 400)
        {
            for (int i = 0; i < maxAttempts; i++)
            {
                bool ok = await action();
                if (ok) return true;
                if (i < maxAttempts - 1) await Task.Delay(retryDelayMs);
            }
            return false;
        }

        private string GetCellType(string columnHeader)
        {
            if (_columnCellTypes == null) return "text";
            return _columnCellTypes.TryGetValue(columnHeader, out string t) ? t ?? "text" : "text";
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
