using System.IO;

namespace AutoFiller.UIDriver
{
    /// <summary>
    /// Configurable delay values (in milliseconds) used throughout the UI
    /// automation layer.  Allows tuning for slow or fast machines without
    /// recompiling.  Load a custom profile via <see cref="LoadOrDefault"/>.
    /// </summary>
    public class TimingConfig
    {
        /// <summary>Process-wide default instance (uses all built-in values).</summary>
        public static TimingConfig Default { get; } = new TimingConfig();

        // ── VisionInteractor timings ───────────────────────────────────────

        /// <summary>Wait after a left-click before continuing (ms).</summary>
        public int ClickDelayMs { get; set; } = 80;

        /// <summary>Wait before sending the click input, after the window is
        /// foregrounded (ms).</summary>
        public int ClickSettleMs { get; set; } = 50;

        /// <summary>Wait for the clipboard API to settle after SetText (ms).</summary>
        public int ClipboardSettleMs { get; set; } = 30;

        /// <summary>Wait after Ctrl+V paste before proceeding (ms).</summary>
        public int AfterPasteMs { get; set; } = 60;

        /// <summary>Wait after clicking a tab header to let it render (ms).</summary>
        public int AfterTabClickMs { get; set; } = 500;

        /// <summary>Wait after posting a function-key message (ms).</summary>
        public int AfterFunctionKeyMs { get; set; } = 30;

        /// <summary>Wait after sending a scroll-wheel event (ms).</summary>
        public int AfterScrollMs { get; set; } = 80;

        /// <summary>Wait after calling <c>SetForegroundWindow</c> (ms).</summary>
        public int BringToFrontMs { get; set; } = 80;

        /// <summary>Wait after sending Ctrl+A (ms).</summary>
        public int SelectAllMs { get; set; } = 30;

        // ── FillOrchestrator timings ───────────────────────────────────────

        /// <summary>Delay between filling individual cells (ms).</summary>
        public int PerCellMs { get; set; } = 120;

        /// <summary>Delay between filling rows (ms).</summary>
        public int PerRowMs { get; set; } = 200;

        /// <summary>Delay after pressing F2 (新規) before continuing (ms).</summary>
        public int AfterF2Ms { get; set; } = 600;

        /// <summary>Polling interval for the confirmation-dialog OCR check (ms).</summary>
        public int VerifyPollMs { get; set; } = 300;

        // ── Optional JSON override ─────────────────────────────────────────

        /// <summary>
        /// Loads timing configuration from a JSON file at <paramref name="jsonPath"/>.
        /// Returns <see cref="Default"/> when the file is absent or unparseable.
        /// </summary>
        public static TimingConfig LoadOrDefault(string jsonPath)
        {
            if (!File.Exists(jsonPath)) return Default;
            try
            {
                string json = File.ReadAllText(jsonPath);
                return System.Text.Json.JsonSerializer.Deserialize<TimingConfig>(json) ?? Default;
            }
            catch
            {
                return Default;
            }
        }
    }
}
