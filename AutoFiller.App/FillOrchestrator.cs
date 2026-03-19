using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using AutoFiller.Core;
using AutoFiller.UIDriver;
using Microsoft.Extensions.Logging;

namespace AutoFiller.App
{
    // ─────────────────────────────────────────────
    // Progress / result types
    // ─────────────────────────────────────────────

    public class FillProgress
    {
        public string Message { get; }
        public bool   Success { get; }
        public FillProgress(string message, bool success = true) { Message = message; Success = success; }
    }

    public class FillResult
    {
        public bool                AutoSubmitted { get; }
        public int                 FilledCells   { get; }
        public int                 ErrorCount    { get; }
        public IReadOnlyList<string> Warnings    { get; }
        public string              FinalMessage  { get; }
        public bool                HasErrors     => ErrorCount > 0;

        public FillResult(
            bool autoSubmitted, int filledCells, int errorCount,
            IReadOnlyList<string> warnings, string finalMessage)
        {
            AutoSubmitted = autoSubmitted;
            FilledCells   = filledCells;
            ErrorCount    = errorCount;
            Warnings      = warnings;
            FinalMessage  = finalMessage;
        }
    }

    // ─────────────────────────────────────────────
    // Orchestrator
    // ─────────────────────────────────────────────

    public class FillOrchestrator
    {
        // ── Virtual key codes ─────────────────────────────────────────────
        private const byte VK_F2     = 0x71;
        private const byte VK_F9     = 0x78;
        private const byte VK_RETURN = 0x0D;

        // ── Dependencies ──────────────────────────────────────────────────
        private readonly VisionInteractor    _vision;
        private readonly TabNavigator        _tabs;
        private readonly ExcelValueExtractor _excel;
        private readonly ScreenCapture       _screen;
        private readonly OcrEngine           _ocr;
        private readonly TimingConfig        _timing;
        private readonly ILogger<FillOrchestrator> _logger;
        private IntPtr                       _hwnd = IntPtr.Zero;

        // ── Run mode ─────────────────────────────────────────────────────
        public enum RunMode { Manual, Auto }

        // ── Constructor ───────────────────────────────────────────────────

        public FillOrchestrator(
            VisionInteractor    vision,
            TabNavigator        tabs,
            ExcelValueExtractor excel,
            TimingConfig        timing = null,
            ILogger<FillOrchestrator> logger = null)
        {
            _vision = vision ?? throw new ArgumentNullException(nameof(vision));
            _tabs   = tabs   ?? throw new ArgumentNullException(nameof(tabs));
            _excel  = excel  ?? throw new ArgumentNullException(nameof(excel));
            _screen = new ScreenCapture();
            _ocr    = new OcrEngine();
            _timing = timing ?? TimingConfig.Default;
            _logger = logger ?? NullLogger<FillOrchestrator>.Instance;
        }

        // ── Attach ────────────────────────────────────────────────────────

        /// <summary>
        /// Attaches to the target app window by partial title match.
        /// Must be called before <see cref="RunAsync"/>.
        /// Returns false when no matching window is found.
        /// </summary>
        public bool Attach(string windowTitleContains)
        {
            bool ok = _vision.Attach(windowTitleContains);
            if (ok)
            {
                _hwnd = _screen.FindWindowByTitle(windowTitleContains);
                _tabs.SetWindowHandle(_hwnd);
            }
            return ok;
        }

        // ── Main entry point ──────────────────────────────────────────────

        /// <summary>
        /// Reads <paramref name="excelFilePath"/>, fills all header fields and
        /// grid rows in the attached app window using the confirmed
        /// <paramref name="mapping"/>, then either returns for manual F9
        /// submission (<see cref="RunMode.Manual"/>) or auto-submits and waits
        /// for the OCR-detected confirmation dialog (<see cref="RunMode.Auto"/>).
        /// </summary>
        public async Task<FillResult> RunAsync(
            string            excelFilePath,
            string            sheetName,
            AppTabStructure   appStructure,
            MappingConfig     mapping,
            RunMode           mode,
            IProgress<FillProgress> progress = null,
            CancellationToken ct             = default)
        {
            if (_hwnd == IntPtr.Zero)
                throw new InvalidOperationException("Call Attach() before RunAsync().");

            // ── STEP 1: Parse Excel ───────────────────────────────────────
            var headerCells = _excel.ExtractHeaderValues(excelFilePath, sheetName);
            var items = _excel.ExtractItemRows(excelFilePath, sheetName)
                              .Where(i => !IsHidden(i))
                              .ToList();

            Report(progress, $"Excel parsed. {headerCells.Count} header cells, {items.Count} item rows.", true);

            // ── STEP 2: Press F2 (新規) to start a new record ─────────────
            _vision.SendFunctionKey(VK_F2);
            await WaitForNewRecordForm(ct);

            // ── STEP 3: Fill header tab ───────────────────────────────────
            await _tabs.NavigateTo(appStructure.HeaderTabName, timeoutMs: 3000);

            foreach (var hm in mapping.HeaderFields)
            {
                ct.ThrowIfCancellationRequested();
                string value = GetHeaderValue(headerCells, hm.ExcelCellAddress);
                if (string.IsNullOrEmpty(value)) continue;

                bool ok = await _vision.FillFieldByLabel(hm.ControlName, value);
                if (!ok)
                {
                    _logger.LogWarning("Header field '{Field}' not found after 3 attempts", hm.ControlName);
                    warnings.Add($"Header field '{hm.ControlName}' not found after 3 attempts");
                }
                Report(progress, $"Header: {hm.ControlName} = {value}", ok);
                await Task.Delay(_timing.PerCellMs, ct);
            }

            // ── STEP 4: Fill grid tab rows ────────────────────────────────
            await _tabs.NavigateTo(appStructure.GridTabName, timeoutMs: 3000);
            await _vision.DetectGrid();   // prime the internal grid cache

            // Refresh column screen positions in case window moved since mapping was saved.
            await _vision.RefreshGridColumnPositions(mapping);

            // Provide cell-type hints ("dropdown" vs. "text") to VisionInteractor.
            if (mapping?.Grid?.Columns != null)
            {
                var types = mapping.Grid.Columns
                    .ToDictionary(
                        kv => kv.Key,
                        kv => kv.Value.CellType ?? "text",
                        StringComparer.OrdinalIgnoreCase);
                _vision.SetColumnCellTypes(types);
            }
            int filled   = 0;
            int errors   = 0;
            var warnings = new List<string>();

            foreach (var item in items)
            {
                ct.ThrowIfCancellationRequested();
                Report(progress, $"Grid row {rowIndex}: {GetItemName(item)}", true);

                // Scroll when the row has scrolled out of the visible grid window.
                if (rowIndex > 0 && rowIndex % 12 == 0)
                    await _vision.ScrollGrid(12);

                foreach (var col in mapping.Grid.Columns)
                {
                    string value = GetItemValue(item, col.Key);
                    if (string.IsNullOrEmpty(value) || value == "－") continue;

                    try
                    {
                        await _vision.FillGridCell(rowIndex, col.Key, value);
                        await Task.Delay(_timing.PerCellMs, ct);
                        filled++;
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Row {Row} col '{Col}' fill failed", rowIndex, col.Key);
                        warnings.Add($"Row {rowIndex} col {col.Key}: {ex.Message}");
                        errors++;
                    }
                }

                rowIndex++;
            }

            // ── STEP 5: Hybrid mode ───────────────────────────────────────
            if (mode == RunMode.Manual)
            {
                return new FillResult(false, filled, errors, warnings,
                    "Fill complete. Press F9 (登録) to submit.");
            }

            _vision.SendFunctionKey(VK_F9);
            string confirmMsg = await WaitForConfirmationDialog();
            return new FillResult(true, filled, errors, warnings, confirmMsg);
        }

        // ── New-record form detection ──────────────────────────────────────

        /// <summary>
        /// After pressing F2 (新規), polls OCR every 300 ms until the form is
        /// ready: two consecutive screenshots with the same word count indicate
        /// the UI has stabilised. Falls through on timeout after
        /// <paramref name="timeoutMs"/> ms.
        /// </summary>
        private async Task WaitForNewRecordForm(CancellationToken ct, int timeoutMs = 5000)
        {
            int elapsed       = 0;
            int prevWordCount = -1;

            while (elapsed < timeoutMs)
            {
                await Task.Delay(300, ct);
                elapsed += 300;

                using var bmp  = _screen.CaptureWindow(_hwnd);
                OcrResult ocr  = await _ocr.RecognizeAsync(bmp);

                // Stable word count → form has finished loading.
                if (ocr.Words.Count == prevWordCount) return;
                prevWordCount = ocr.Words.Count;
            }
            // Timeout: proceed anyway.
        }

        // ── Confirmation polling ──────────────────────────────────────────

        /// <summary>
        /// Takes OCR screenshots every 300 ms until "登録", "完了", or
        /// "しました" appears (final confirmation dialog), then dismisses it
        /// with Enter. Returns the matched word text, or a timeout message.
        /// </summary>
        private async Task<string> WaitForConfirmationDialog(int timeoutMs = 10000)
        {
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                using Bitmap bmp = _screen.CaptureWindow(_hwnd);
                OcrResult ocr   = await _ocr.RecognizeAsync(bmp);

                OcrWord confirm = ocr.Words.FirstOrDefault(w =>
                    w.Text.Contains("登録")     ||
                    w.Text.Contains("完了")     ||
                    w.Text.Contains("しました"));

                if (confirm != null)
                {
                    _vision.SendFunctionKey(VK_RETURN);
                    return confirm.Text;
                }

                await Task.Delay(_timing.VerifyPollMs);
            }

            return "Timeout waiting for confirmation";
        }

        // ── Private helpers ───────────────────────────────────────────────

        /// <summary>
        /// Looks up a header cell by its Excel address (e.g. "B5") and
        /// returns its raw value, or an empty string when not found.
        /// </summary>
        private static string GetHeaderValue(
            IReadOnlyList<ExcelCellValue> headerCells,
            string cellAddress)
        {
            if (string.IsNullOrEmpty(cellAddress)) return string.Empty;

            ExcelCellValue cell = headerCells.FirstOrDefault(c =>
                string.Equals(c.CellAddress, cellAddress, StringComparison.OrdinalIgnoreCase));

            return cell?.RawValue ?? string.Empty;
        }

        /// <summary>
        /// Returns the item row value for <paramref name="columnHeader"/>,
        /// or an empty string when the column is absent in this row.
        /// </summary>
        private static string GetItemValue(ItemRowValues item, string columnHeader)
        {
            return item.Values.TryGetValue(columnHeader, out string v) ? v : string.Empty;
        }

        /// <summary>
        /// Returns the 項目名 value for progress display, falling back to the
        /// row's Excel source row number when 項目名 is absent or blank.
        /// </summary>
        private static string GetItemName(ItemRowValues item)
        {
            if (item.Values.TryGetValue("項目名", out string name) && !string.IsNullOrEmpty(name))
                return name;
            return $"行 {item.SourceRow}";
        }

        /// <summary>
        /// Returns true when the row's 変更区分 column value is "非表示".
        /// Rows without that column are treated as visible.
        /// </summary>
        private static bool IsHidden(ItemRowValues item)
        {
            return item.Values.TryGetValue("変更区分", out string v) && v == "非表示";
        }

        private static void Report(IProgress<FillProgress> progress, string message, bool success)
        {
            progress?.Report(new FillProgress(message, success));
        }
    }
}
