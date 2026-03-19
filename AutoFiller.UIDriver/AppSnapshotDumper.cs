using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace AutoFiller.UIDriver
{
    // ─────────────────────────────────────────────
    // Data model
    // ─────────────────────────────────────────────

    public class LabelValuePair
    {
        public string LabelText { get; set; }
        public string ValueText { get; set; }
        /// <summary>Absolute screen X of the value field centre.</summary>
        public int ClickX { get; set; }
        /// <summary>Absolute screen Y of the value field centre.</summary>
        public int ClickY { get; set; }
    }

    public class GridCellSnapshot
    {
        public string ColumnHeader { get; set; }
        public string Value { get; set; }
        public int X { get; set; }
        public int Y { get; set; }
    }

    public class GridRowSnapshot
    {
        public int RowIndex { get; set; }
        public int Y { get; set; }
        public List<GridCellSnapshot> Cells { get; set; } = new List<GridCellSnapshot>();
    }

    public class GridVisionSnapshot
    {
        public int HeaderRowY { get; set; }
        public int FirstDataRowY { get; set; }
        public int RowHeight { get; set; }
        public List<string> ColumnHeaders { get; set; } = new List<string>();
        public List<GridRowSnapshot> Rows { get; set; } = new List<GridRowSnapshot>();
    }

    public class TabVisionSnapshot
    {
        public string TabName { get; set; }
        [JsonIgnore]
        public Bitmap Screenshot { get; set; }
        [JsonIgnore]
        public OcrResult OcrData { get; set; }
        public List<LabelValuePair> LabeledFields { get; set; } = new List<LabelValuePair>();
        public GridVisionSnapshot Grid { get; set; }
    }

    public class VisionSnapshot
    {
        public string WindowTitle { get; set; }
        public DateTime CapturedAt { get; set; }
        public List<TabVisionSnapshot> Tabs { get; set; } = new List<TabVisionSnapshot>();
    }

    // ─────────────────────────────────────────────
    // Dumper
    // ─────────────────────────────────────────────

    public class VisionSnapshotDumper
    {
        // ── tuning constants ──────────────────────────────────────────────

        /// <summary>Max distance (px) to the right of a label to look for a value word.</summary>
        private const int MaxRightValueDistance = 300;
        /// <summary>Max distance (px) below a label to look for a value word.</summary>
        private const int MaxBelowValueDistance = 40;
        /// <summary>Vertical tolerance (px) for treating two words as on the same row.</summary>
        private const int YBandTolerance = 8;
        /// <summary>Horizontal tolerance (px) for treating two words as in the same column.</summary>
        private const int XBandTolerance = 50;
        /// <summary>Minimum character length for a token to be treated as a label.</summary>
        private const int MinLabelLength = 2;

        // ── fields ────────────────────────────────────────────────────────

        private readonly OcrEngine _ocr;
        private readonly ScreenCapture _capture;

        // ── constructors ──────────────────────────────────────────────────

        public VisionSnapshotDumper() : this(OcrEngineProvider.Instance) { }

        public VisionSnapshotDumper(OcrEngine ocrEngine)
        {
            _ocr     = ocrEngine ?? throw new ArgumentNullException(nameof(ocrEngine));
            _capture = new ScreenCapture();
        }

        // ── public API ────────────────────────────────────────────────────

        /// <summary>
        /// Navigates every tab of the window identified by <paramref name="hwnd"/>,
        /// captures a screenshot of each, runs OCR, and collects label-value pairs
        /// and grid data.
        /// </summary>
        public async Task<VisionSnapshot> DumpFullSnapshot(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero)
                throw new ArgumentException("hwnd must not be Zero.", nameof(hwnd));

            var snapshot = new VisionSnapshot
            {
                WindowTitle = GetWindowTitle(hwnd),
                CapturedAt  = DateTime.Now
            };

            // Build VisionInteractor + TabNavigator, both bound to the same hwnd.
            var interactor = new VisionInteractor(_ocr);
            interactor.Attach(snapshot.WindowTitle);   // resolves hwnd from title
            var navigator = new TabNavigator(interactor, _ocr);
            navigator.SetWindowHandle(hwnd);           // ensure exact hwnd is used

            List<string> tabNames = await navigator.DiscoverTabs();

            if (tabNames.Count == 0)
            {
                // No tab strip — snapshot the single view.
                snapshot.Tabs.Add(await CaptureSingleTab(hwnd, "(No Tabs)", interactor));
            }
            else
            {
                foreach (string tabName in tabNames)
                {
                    bool navigated = await navigator.NavigateTo(tabName, timeoutMs: 3000);
                    if (!navigated)
                        continue;
                    snapshot.Tabs.Add(await CaptureSingleTab(hwnd, tabName, interactor));
                }
            }

            return snapshot;
        }

        /// <summary>Serialises the snapshot to an indented JSON file.</summary>
        /// <remarks>
        /// <see cref="Bitmap"/> and <see cref="OcrResult"/> properties are excluded
        /// via <c>[JsonIgnore]</c> on <see cref="TabVisionSnapshot"/>.
        /// </remarks>
        public void ExportToJson(VisionSnapshot snapshot, string outputPath)
        {
            if (snapshot    == null) throw new ArgumentNullException(nameof(snapshot));
            if (string.IsNullOrEmpty(outputPath)) throw new ArgumentException("outputPath must not be empty.", nameof(outputPath));

            Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };

            File.WriteAllText(outputPath, JsonSerializer.Serialize(snapshot, options), Encoding.UTF8);
        }

        /// <summary>
        /// Saves a flat CSV with columns: TabName, LabelText, ValueText, ClickX, ClickY.
        /// One row per <see cref="LabelValuePair"/> across all tabs.
        /// </summary>
        public void ExportToCsv(VisionSnapshot snapshot, string outputPath)
        {
            if (snapshot    == null) throw new ArgumentNullException(nameof(snapshot));
            if (string.IsNullOrEmpty(outputPath)) throw new ArgumentException("outputPath must not be empty.", nameof(outputPath));

            Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

            var sb = new StringBuilder();
            sb.AppendLine("TabName,LabelText,ValueText,ClickX,ClickY");

            foreach (var tab in snapshot.Tabs)
            {
                foreach (var pair in tab.LabeledFields)
                {
                    sb.AppendLine(string.Join(",",
                        CsvEscape(tab.TabName   ?? string.Empty),
                        CsvEscape(pair.LabelText ?? string.Empty),
                        CsvEscape(pair.ValueText ?? string.Empty),
                        pair.ClickX,
                        pair.ClickY));
                }
            }

            File.WriteAllText(outputPath, sb.ToString(), Encoding.UTF8);
        }

        /// <summary>
        /// Returns a safe desktop path:
        /// <c>Desktop\snapshot_{title}_{yyyyMMdd_HHmmss}.json</c>
        /// </summary>
        public static string DefaultOutputPath(string windowTitle, string extension = "json")
        {
            string safe = string.Concat(
                (windowTitle ?? "app")
                .Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '_' : c));

            string ts      = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            return Path.Combine(desktop, $"snapshot_{safe}_{ts}.{extension}");
        }

        // ── private: per-tab capture ──────────────────────────────────────

        private async Task<TabVisionSnapshot> CaptureSingleTab(
            IntPtr hwnd, string tabName, VisionInteractor interactor)
        {
            Bitmap    screenshot  = _capture.CaptureWindow(hwnd);
            Rectangle windowBounds = _capture.GetWindowBounds(hwnd);
            OcrResult ocrResult  = await _ocr.RecognizeAsync(screenshot);

            GridLayout?         gridLayout   = await interactor.DetectGrid();
            GridVisionSnapshot? gridSnapshot = gridLayout != null
                ? DumpGrid(ocrResult, gridLayout, windowBounds)
                : null;

            return new TabVisionSnapshot
            {
                TabName       = tabName,
                Screenshot    = screenshot,
                OcrData       = ocrResult,
                LabeledFields = DetectLabelValuePairs(ocrResult, windowBounds),
                Grid          = gridSnapshot
            };
        }

        // ── private: label-value detection ───────────────────────────────

        /// <summary>
        /// Scans OCR words for adjacent label→value pairs using two heuristics:
        /// <list type="bullet">
        ///   <item><b>RIGHT</b>: value word lies within <see cref="MaxRightValueDistance"/> px
        ///         to the right and within <see cref="YBandTolerance"/> px vertically.</item>
        ///   <item><b>BELOW</b>: value word lies within <see cref="MaxBelowValueDistance"/> px
        ///         below and within <see cref="XBandTolerance"/> px horizontally.</item>
        /// </list>
        /// Returned <c>ClickX/Y</c> are absolute screen coordinates
        /// (bitmap offset + window origin).
        /// </summary>
        private List<LabelValuePair> DetectLabelValuePairs(OcrResult ocr, Rectangle windowBounds)
        {
            var pairs = new List<LabelValuePair>();
            if (ocr?.Words == null || ocr.Words.Count == 0) return pairs;

            var words       = ocr.Words.Where(w => !string.IsNullOrWhiteSpace(w.Text)).ToList();
            var usedAsValue = new HashSet<OcrWord>(ReferenceEqualityComparer.Instance);

            foreach (var label in words)
            {
                if (label.Text.Length < MinLabelLength) continue;
                if (usedAsValue.Contains(label))        continue;

                int lRight = label.BoundingBox.Right;
                int lMidY  = label.BoundingBox.Top + label.BoundingBox.Height / 2;
                int lCx    = label.BoundingBox.Left + label.BoundingBox.Width  / 2;
                int lBottom = label.BoundingBox.Bottom;

                // ── RIGHT heuristic ───────────────────────────────────────
                OcrWord? rightCand = words
                    .Where(w => w != label
                             && !usedAsValue.Contains(w)
                             && w.BoundingBox.Left > lRight
                             && w.BoundingBox.Left - lRight <= MaxRightValueDistance
                             && Math.Abs(w.BoundingBox.Top + w.BoundingBox.Height / 2 - lMidY) <= YBandTolerance)
                    .OrderBy(w => w.BoundingBox.Left)
                    .FirstOrDefault();

                if (rightCand != null)
                {
                    pairs.Add(MakePair(label.Text, rightCand, windowBounds));
                    usedAsValue.Add(rightCand);
                    continue;
                }

                // ── BELOW heuristic ───────────────────────────────────────
                OcrWord? belowCand = words
                    .Where(w => w != label
                             && !usedAsValue.Contains(w)
                             && w.BoundingBox.Top > lBottom
                             && w.BoundingBox.Top - lBottom <= MaxBelowValueDistance
                             && Math.Abs(w.BoundingBox.Left + w.BoundingBox.Width / 2 - lCx) <= XBandTolerance)
                    .OrderBy(w => w.BoundingBox.Top)
                    .FirstOrDefault();

                if (belowCand != null)
                {
                    pairs.Add(MakePair(label.Text, belowCand, windowBounds));
                    usedAsValue.Add(belowCand);
                }
            }

            return pairs;
        }

        private static LabelValuePair MakePair(string labelText, OcrWord value, Rectangle windowBounds)
        {
            int bmpCx = value.BoundingBox.Left + value.BoundingBox.Width  / 2;
            int bmpCy = value.BoundingBox.Top  + value.BoundingBox.Height / 2;
            return new LabelValuePair
            {
                LabelText = labelText,
                ValueText = value.Text,
                ClickX    = windowBounds.X + bmpCx,
                ClickY    = windowBounds.Y + bmpCy
            };
        }

        // ── private: grid dump ────────────────────────────────────────────

        /// <summary>
        /// Bins OCR words into grid rows using the <paramref name="gridLayout"/>
        /// row geometry (absolute screen coords → bitmap coords via window origin).
        /// Words in the same row+column are concatenated with a space.
        /// </summary>
        private GridVisionSnapshot DumpGrid(
            OcrResult ocrResult, GridLayout gridLayout, Rectangle windowBounds)
        {
            var snap = new GridVisionSnapshot
            {
                HeaderRowY    = gridLayout.HeaderRowY,
                FirstDataRowY = gridLayout.FirstDataRowY,
                RowHeight     = gridLayout.RowHeight,
                ColumnHeaders = gridLayout.Columns.Select(c => c.HeaderText).ToList()
            };

            if (ocrResult?.Words == null || ocrResult.Words.Count == 0) return snap;

            // Convert screen Y thresholds to bitmap Y.
            int bmpFirstDataY = gridLayout.FirstDataRowY - windowBounds.Y;
            int rowH          = Math.Max(gridLayout.RowHeight, 1);
            int halfRow       = rowH / 2;

            // Column centres in bitmap X.
            var bmpCols = gridLayout.Columns
                .Select(c => new { c.HeaderText, BmpCx = c.CenterX - windowBounds.X, c.Width })
                .ToList();

            // Filter to data-row words only.
            var dataWords = ocrResult.Words
                .Where(w => !string.IsNullOrWhiteSpace(w.Text)
                         && w.BoundingBox.Top >= bmpFirstDataY - halfRow)
                .ToList();

            // Group words into row bands.
            var rowGroups = new Dictionary<int, List<OcrWord>>();
            foreach (var word in dataWords)
            {
                int wordCy   = word.BoundingBox.Top + word.BoundingBox.Height / 2;
                int rowIndex = (wordCy - bmpFirstDataY) / rowH;
                if (rowIndex < 0) rowIndex = 0;

                if (!rowGroups.ContainsKey(rowIndex))
                    rowGroups[rowIndex] = new List<OcrWord>();
                rowGroups[rowIndex].Add(word);
            }

            foreach (var kvp in rowGroups.OrderBy(k => k.Key))
            {
                int rowIdx     = kvp.Key;
                int rowScreenY = gridLayout.FirstDataRowY + rowIdx * rowH + halfRow;

                var cells = new List<GridCellSnapshot>();
                foreach (var word in kvp.Value.OrderBy(w => w.BoundingBox.Left))
                {
                    int wordBmpCx = word.BoundingBox.Left + word.BoundingBox.Width / 2;

                    // Assign to nearest column header.
                    var col        = bmpCols.OrderBy(c => Math.Abs(c.BmpCx - wordBmpCx)).FirstOrDefault();
                    string header  = col?.HeaderText ?? string.Empty;
                    int cellScreenX = col != null ? col.BmpCx + windowBounds.X : windowBounds.X + wordBmpCx;

                    // Merge multiple words in the same cell.
                    var existing = cells.FirstOrDefault(c => c.ColumnHeader == header);
                    if (existing != null)
                        existing.Value = (existing.Value + " " + word.Text).Trim();
                    else
                        cells.Add(new GridCellSnapshot
                        {
                            ColumnHeader = header,
                            Value        = word.Text,
                            X            = cellScreenX,
                            Y            = rowScreenY
                        });
                }

                snap.Rows.Add(new GridRowSnapshot { RowIndex = rowIdx, Y = rowScreenY, Cells = cells });
            }

            return snap;
        }

        // ── utility ───────────────────────────────────────────────────────

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        private static string GetWindowTitle(IntPtr hwnd)
        {
            var sb = new StringBuilder(512);
            GetWindowText(hwnd, sb, sb.Capacity);
            return sb.ToString();
        }

        private static string CsvEscape(string value)
        {
            if (value.Contains(',') || value.Contains('"') || value.Contains('\n'))
                return $"\"{value.Replace("\"", "\"\"")}\"";
            return value;
        }
    }
}
