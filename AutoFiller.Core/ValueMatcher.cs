// References types from:
//   AutoFiller.UIDriver  (VisionSnapshot, TabVisionSnapshot, LabelValuePair,
//                         GridVisionSnapshot, GridRowSnapshot, GridCellSnapshot)
//   AutoFiller.Core      (ExcelCellValue, ItemRowValues, ExcelValueExtractor)

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using AutoFiller.UIDriver;

namespace AutoFiller.Core
{
    // ─────────────────────────────────────────────
    // Result model
    // ─────────────────────────────────────────────

    public enum MatchType
    {
        ExactMatch,
        PartialMatch,
        GridCellMatch,
        NoMatch
    }

    public class MatchResult
    {
        public string ControlName { get; set; }
        public string TabContext { get; set; }
        public string ControlPath { get; set; }
        public double ControlX { get; set; }
        public double ControlY { get; set; }
        public string ExcelSheet { get; set; }
        public string ExcelCellAddress { get; set; }
        public int ExcelRow { get; set; }
        public int ExcelCol { get; set; }
        public string MatchedValue { get; set; }
        public double ConfidenceScore { get; set; }
        public MatchType Type { get; set; }
    }

    public class HeaderFieldMapping
    {
        public string ControlName { get; set; }
        public string TabContext { get; set; }
        public double ClickX { get; set; }
        public double ClickY { get; set; }
        public string ExcelCellAddress { get; set; }
        public int ExcelRow { get; set; }
        public int ExcelCol { get; set; }
        /// <summary>"Clipboard", "SendKeys", or "Dropdown"</summary>
        public string InputMethod { get; set; }
    }

    public class GridColumnMapping
    {
        public string GridColumnHeader { get; set; }
        public int ExcelColIndex { get; set; }
        public string ExcelColLetter { get; set; }
        /// <summary>
        /// Volatile screen X-coordinate of the column centre. Not persisted to JSON
        /// (depends on current window position). Repopulate via
        /// <c>VisionInteractor.RefreshGridColumnPositions()</c> after loading from JSON.
        /// </summary>
        [System.Text.Json.Serialization.JsonIgnore]
        public double GridColumnX { get; set; }
        public string CellType { get; set; }
    }

    public class GridMapping
    {
        public double GridOriginX { get; set; }
        public double GridOriginY { get; set; }
        public double RowHeight { get; set; }
        public string TabContext { get; set; }
        public Dictionary<string, GridColumnMapping> Columns { get; set; }
            = new Dictionary<string, GridColumnMapping>();
    }

    public class UnmatchedControl
    {
        public string ControlName { get; set; }
        public string TabContext { get; set; }
        public string CurrentValue { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
    }

    public class MappingConfig
    {
        public string AppWindowTitle { get; set; }
        public string ExcelFilePath { get; set; }
        public DateTime GeneratedAt { get; set; }
        public List<HeaderFieldMapping> HeaderFields { get; set; } = new();
        public GridMapping Grid { get; set; }
        public List<UnmatchedControl> UnmatchedControls { get; set; } = new();
        public List<string> UnmatchedExcelValues { get; set; } = new();
    }

    // ─────────────────────────────────────────────
    // Matcher
    // ─────────────────────────────────────────────

    public class ValueMatcher
    {
        // Values shorter than this are only matchable if they look like numeric IDs.
        private const int MinMatchableLength = 3;

        // Cells whose normalised value appears more than this many times in Excel
        // are considered non-unique and skipped.
        private const int MaxExcelOccurrences = 3;

        private readonly ExcelValueExtractor _extractor = new ExcelValueExtractor();

        // ── public API ────────────────────────────────────────────────────────

        /// <summary>
        /// Compares every non-empty control value in <paramref name="snapshot"/>
        /// against <paramref name="excelValues"/> and produces a
        /// <see cref="MappingConfig"/> describing the discovered mappings.
        /// </summary>
        public MappingConfig Match(VisionSnapshot snapshot, List<ExcelCellValue> excelValues)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            if (excelValues == null) throw new ArgumentNullException(nameof(excelValues));

            // ── Step 1: Build normalised Excel lookup ──────────────────────
            // normalizedValue → list of matching ExcelCellValue entries.
            var excelByValue = BuildExcelLookup(excelValues);

            // Pre-compute per-value occurrence counts for uniqueness filtering.
            var occurrenceCounts = excelValues
                .GroupBy(c => c.NormalizedValue ?? string.Empty)
                .ToDictionary(g => g.Key, g => g.Count());

            // ── Step 2 & 4: Match header controls (non-grid) ──────────────
            var headerMatches = new List<MatchResult>();
            var matchedExcelAddresses = new HashSet<string>(StringComparer.Ordinal);

            // Build a flat list of all label-value pairs with their tab context.
            // We distinguish grid vs. non-grid by TabVisionSnapshot.Grid presence.
            var allLabelValues = snapshot.Tabs
                .SelectMany(t => t.LabeledFields
                    .Select(p => (TabName: t.TabName ?? string.Empty, Pair: p)))
                .ToList();

            var gridControlNames = BuildGridControlNameSet(snapshot);

            foreach (var entry in allLabelValues)
            {
                string rawValue = entry.Pair.ValueText ?? string.Empty;
                if (string.IsNullOrWhiteSpace(rawValue)) continue;

                // Skip label-value pairs whose label appears in a detected grid.
                if (gridControlNames.Contains(entry.Pair.LabelText ?? string.Empty)) continue;

                string norm = _extractor.Normalize(rawValue);
                if (!IsMatchableValue(norm, occurrenceCounts)) continue;

                if (!excelByValue.TryGetValue(norm, out var candidates)) continue;

                // Score: exact = 1.0; prefer HEADER section cells.
                var best = candidates
                    .OrderByDescending(c => c.SectionHint == "HEADER" ? 1 : 0)
                    .ThenBy(c => c.Row)
                    .First();

                double confidence = 1.0;

                // Penalise if duplicate Excel cell already used for another control.
                if (matchedExcelAddresses.Contains(best.CellAddress))
                    confidence = 0.7;

                matchedExcelAddresses.Add(best.CellAddress);

                headerMatches.Add(new MatchResult
                {
                    ControlName      = entry.Pair.LabelText ?? string.Empty,
                    TabContext        = entry.TabName,
                    ControlPath      = entry.Pair.LabelText ?? string.Empty,
                    ControlX         = entry.Pair.ClickX,
                    ControlY         = entry.Pair.ClickY,
                    ExcelSheet       = best.SheetName,
                    ExcelCellAddress = best.CellAddress,
                    ExcelRow         = best.Row,
                    ExcelCol         = best.Col,
                    MatchedValue     = rawValue,
                    ConfidenceScore  = confidence,
                    Type             = MatchType.ExactMatch
                });
            }

            // ── Step 3: Grid mapping ───────────────────────────────────────
            var excelItemRows = excelValues
                .Where(c => c.SectionHint == "ITEM_TABLE")
                .GroupBy(c => c.Row)
                .Select(g => new ItemRowValues
                {
                    SourceRow = g.Key,
                    Values = g.ToDictionary(
                        c => c.ColLetter,
                        c => c.RawValue ?? string.Empty)
                })
                .OrderBy(r => r.SourceRow)
                .ToList();

            GridMapping gridMapping = BuildGridMapping(snapshot, excelItemRows, excelValues);

            // ── Step 5: Build MappingConfig ────────────────────────────────
            var headerFields = headerMatches
                .Where(m => m.ConfidenceScore > 0.5)
                .Select(m => new HeaderFieldMapping
                {
                    ControlName = m.ControlName,
                    TabContext = m.TabContext,
                    ClickX = m.ControlX,
                    ClickY = m.ControlY,
                    ExcelCellAddress = m.ExcelCellAddress,
                    ExcelRow = m.ExcelRow,
                    ExcelCol = m.ExcelCol,
                    InputMethod = InferInputMethod(m.MatchedValue)
                })
                .ToList();

            var matchedControlNames = new HashSet<string>(
                headerFields.Select(f => f.ControlName), StringComparer.Ordinal);

            var unmatchedControls = allLabelValues
                .Where(e => !string.IsNullOrWhiteSpace(e.Pair.ValueText)
                            && !gridControlNames.Contains(e.Pair.LabelText ?? string.Empty)
                            && !matchedControlNames.Contains(e.Pair.LabelText ?? string.Empty))
                .Select(e => new UnmatchedControl
                {
                    ControlName  = e.Pair.LabelText ?? string.Empty,
                    TabContext    = e.TabName,
                    CurrentValue = e.Pair.ValueText,
                    X            = e.Pair.ClickX,
                    Y            = e.Pair.ClickY
                })
                .ToList();

            var usedAddresses = new HashSet<string>(
                headerFields.Select(f => f.ExcelCellAddress), StringComparer.Ordinal);

            var unmatchedExcel = excelValues
                .Where(c => !usedAddresses.Contains(c.CellAddress)
                            && !string.IsNullOrWhiteSpace(c.NormalizedValue))
                .Select(c => $"{c.CellAddress}: {c.RawValue}")
                .ToList();

            return new MappingConfig
            {
                AppWindowTitle = snapshot.WindowTitle ?? string.Empty,
                ExcelFilePath = string.Empty,   // caller sets this
                GeneratedAt = DateTime.Now,
                HeaderFields = headerFields,
                Grid = gridMapping,
                UnmatchedControls = unmatchedControls,
                UnmatchedExcelValues = unmatchedExcel
            };
        }

        /// <summary>
        /// Serialises <paramref name="config"/> to an indented JSON file.
        /// </summary>
        public void SaveMappingConfig(MappingConfig config, string outputPath)
        {
            if (config == null) throw new ArgumentNullException(nameof(config));
            if (string.IsNullOrEmpty(outputPath))
                throw new ArgumentException("outputPath must not be empty.", nameof(outputPath));

            Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };

            File.WriteAllText(outputPath, JsonSerializer.Serialize(config, options), Encoding.UTF8);
        }

        /// <summary>
        /// Generates a human-readable Markdown report of the mapping results
        /// and writes it to <paramref name="outputPath"/>.
        /// </summary>
        public void SaveMappingReport(MappingConfig config, string outputPath)
        {
            if (config == null) throw new ArgumentNullException(nameof(config));
            if (string.IsNullOrEmpty(outputPath))
                throw new ArgumentException("outputPath must not be empty.", nameof(outputPath));

            Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

            var sb = new StringBuilder();

            sb.AppendLine("# Mapping Report");
            sb.AppendLine();
            sb.AppendLine($"- **App window**: {config.AppWindowTitle}");
            sb.AppendLine($"- **Excel file**: {config.ExcelFilePath}");
            sb.AppendLine($"- **Generated**: {config.GeneratedAt:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine();

            // ── Header fields ──────────────────────────────────────────────
            sb.AppendLine("## Header Field Mappings");
            sb.AppendLine();
            if (config.HeaderFields?.Count > 0)
            {
                sb.AppendLine("| Control | Tab | Excel Cell | Value | Input Method | Click (X,Y) |");
                sb.AppendLine("|---------|-----|------------|-------|--------------|-------------|");
                foreach (var f in config.HeaderFields.OrderBy(f => f.ExcelRow))
                {
                    sb.AppendLine(
                        $"| {Md(f.ControlName)} | {Md(f.TabContext)} | {Md(f.ExcelCellAddress)} " +
                        $"| | {Md(f.InputMethod)} | ({f.ClickX:F0}, {f.ClickY:F0}) |");
                }
            }
            else
            {
                sb.AppendLine("_No header field mappings found._");
            }
            sb.AppendLine();

            // ── Grid mapping ───────────────────────────────────────────────
            sb.AppendLine("## Grid Column Mappings");
            sb.AppendLine();
            if (config.Grid?.Columns?.Count > 0)
            {
                sb.AppendLine($"- **Tab**: {config.Grid.TabContext}");
                sb.AppendLine($"- **Grid origin**: ({config.Grid.GridOriginX:F0}, {config.Grid.GridOriginY:F0})");
                sb.AppendLine($"- **Row height**: {config.Grid.RowHeight:F1} px");
                sb.AppendLine();
                sb.AppendLine("| Grid Header | Excel Col | Excel Col Letter | Screen X | Cell Type |");
                sb.AppendLine("|-------------|-----------|-----------------|----------|-----------|");
                foreach (var kv in config.Grid.Columns.OrderBy(c => c.Value.ExcelColIndex))
                {
                    var col = kv.Value;
                    sb.AppendLine(
                        $"| {Md(col.GridColumnHeader)} | {col.ExcelColIndex} | {Md(col.ExcelColLetter)} " +
                        $"| {col.GridColumnX:F0} | {Md(col.CellType)} |");
                }
            }
            else
            {
                sb.AppendLine("_No grid mapping found._");
            }
            sb.AppendLine();

            // ── Unmatched controls ─────────────────────────────────────────
            sb.AppendLine("## Unmatched Controls (need manual review)");
            sb.AppendLine();
            if (config.UnmatchedControls?.Count > 0)
            {
                sb.AppendLine("| Control | Tab | Current Value | X | Y |");
                sb.AppendLine("|---------|-----|---------------|---|---|");
                foreach (var u in config.UnmatchedControls)
                    sb.AppendLine(
                        $"| {Md(u.ControlName)} | {Md(u.TabContext)} | {Md(u.CurrentValue)} " +
                        $"| {u.X:F0} | {u.Y:F0} |");
            }
            else
            {
                sb.AppendLine("_All controls were matched._");
            }
            sb.AppendLine();

            // ── Unmatched Excel values ─────────────────────────────────────
            sb.AppendLine("## Unmatched Excel Values (in Excel but not found in app)");
            sb.AppendLine();
            if (config.UnmatchedExcelValues?.Count > 0)
            {
                foreach (var v in config.UnmatchedExcelValues)
                    sb.AppendLine($"- {v}");
            }
            else
            {
                sb.AppendLine("_All Excel values were matched._");
            }

            File.WriteAllText(outputPath, sb.ToString(), Encoding.UTF8);
        }

        // ── private: matching helpers ─────────────────────────────────────────

        /// <summary>
        /// Returns true when <paramref name="normalizedValue"/> is specific enough
        /// to use as a match key:
        /// <list type="bullet">
        ///   <item>At least <see cref="MinMatchableLength"/> characters, OR</item>
        ///   <item>A short purely-numeric / circled-number ID (length 1-2).</item>
        ///   <item>Not appearing more than <see cref="MaxExcelOccurrences"/> times in Excel.</item>
        /// </list>
        /// </summary>
        private bool IsMatchableValue(
            string normalizedValue,
            Dictionary<string, int> occurrenceCounts)
        {
            if (string.IsNullOrWhiteSpace(normalizedValue)) return false;

            // Allow short values only when they look like identifiers.
            if (normalizedValue.Length < MinMatchableLength)
            {
                if (!IsNumericId(normalizedValue)) return false;
            }

            occurrenceCounts.TryGetValue(normalizedValue, out int count);
            return count <= MaxExcelOccurrences;
        }

        private static bool IsNumericId(string value)
        {
            // Accepts ASCII digits and circled numbers ①–⑳ (U+2460–U+2473).
            foreach (char c in value)
            {
                bool isAsciiDigit = c >= '0' && c <= '9';
                bool isCircled = c >= '\u2460' && c <= '\u2473';
                if (!isAsciiDigit && !isCircled) return false;
            }
            return value.Length > 0;
        }

        /// <summary>
        /// Scores an app grid row against all Excel item rows.
        /// Score = (matching cell count) / (total non-empty cells in grid row).
        /// Returns the best-matching Excel source row and its score.
        /// Returns (-1, 0.0) when no plausible match is found.
        /// </summary>
        private (int excelRow, double score) MatchGridRow(
            GridRowSnapshot gridRow,
            List<ItemRowValues> excelItems)
        {
            if (gridRow?.Cells == null || gridRow.Cells.Count == 0)
                return (-1, 0.0);

            var nonEmptyCells = gridRow.Cells
                .Where(c => !string.IsNullOrWhiteSpace(c.Value))
                .ToList();

            if (nonEmptyCells.Count == 0) return (-1, 0.0);

            var normGridValues = nonEmptyCells
                .Select(c => _extractor.Normalize(c.Value))
                .ToHashSet(StringComparer.Ordinal);

            int bestRow = -1;
            double bestScore = 0.0;

            foreach (var excelItem in excelItems)
            {
                int matches = 0;
                foreach (var kv in excelItem.Values)
                {
                    string normExcel = _extractor.Normalize(kv.Value);
                    if (normGridValues.Contains(normExcel)) matches++;
                }

                double score = nonEmptyCells.Count > 0
                    ? (double)matches / nonEmptyCells.Count
                    : 0.0;

                if (score > bestScore)
                {
                    bestScore = score;
                    bestRow = excelItem.SourceRow;
                }
            }

            // Require at least one match.
            return bestScore > 0 ? (bestRow, bestScore) : (-1, 0.0);
        }

        // ── private: grid mapping builder ─────────────────────────────────────

        private GridMapping BuildGridMapping(
            VisionSnapshot snapshot,
            List<ItemRowValues> excelItemRows,
            List<ExcelCellValue> excelValues)
        {
            // Find the first tab that has a detected grid.
            TabVisionSnapshot tabWithGrid = snapshot.Tabs?.FirstOrDefault(t => t.Grid != null);
            if (tabWithGrid == null) return null;

            var grid = tabWithGrid.Grid;

            // Derive per-column X positions from the first data row's cells.
            var colXPositions = grid.Rows.Count > 0
                ? grid.Rows[0].Cells.ToDictionary(
                    c => c.ColumnHeader ?? string.Empty,
                    c => (double)c.X,
                    StringComparer.Ordinal)
                : new Dictionary<string, double>();

            var mapping = new GridMapping
            {
                GridOriginX = colXPositions.Values.FirstOrDefault(),
                GridOriginY = grid.FirstDataRowY,
                RowHeight   = grid.RowHeight,
                TabContext   = tabWithGrid.TabName ?? string.Empty
            };

            if (grid.Rows.Count == 0 || excelItemRows.Count == 0)
                return mapping;

            // ── Match each app grid row against an Excel item row. ─────────
            // Column-header → Excel column index counts (vote-based inference).
            // For each grid column, tally which Excel columns it matched.
            // Dictionary<gridColHeader, Dictionary<excelColLetter, vote count>>
            var colVotes = new Dictionary<string, Dictionary<string, int>>(StringComparer.Ordinal);

            // Pre-build a lookup: normalised value → ExcelCellValue for item rows.
            var excelItemLookup = excelValues
                .Where(c => c.SectionHint == "ITEM_TABLE")
                .GroupBy(c => c.NormalizedValue ?? string.Empty)
                .ToDictionary(g => g.Key, g => g.ToList());

            foreach (var gridRow in grid.Rows)
            {
                var (bestExcelRow, score) = MatchGridRow(gridRow, excelItemRows);
                if (bestExcelRow < 0 || score < 0.3) continue;

                // Build a fast value → column-letter map for this Excel row.
                var excelRowCells = excelValues
                    .Where(c => c.Row == bestExcelRow)
                    .ToDictionary(
                        c => _extractor.Normalize(c.RawValue ?? string.Empty),
                        c => c.ColLetter,
                        StringComparer.Ordinal);

                foreach (var cell in gridRow.Cells)
                {
                    if (string.IsNullOrWhiteSpace(cell.Value)) continue;

                    string normCell = _extractor.Normalize(cell.Value);
                    if (!excelRowCells.TryGetValue(normCell, out string excelLetter)) continue;

                    string gridHdr = cell.ColumnHeader ?? string.Empty;
                    if (!colVotes.TryGetValue(gridHdr, out var votes))
                    {
                        votes = new Dictionary<string, int>(StringComparer.Ordinal);
                        colVotes[gridHdr] = votes;
                    }

                    votes.TryGetValue(excelLetter, out int prev);
                    votes[excelLetter] = prev + 1;
                }
            }

            // Promote the highest-voted Excel column for each grid column header.
            foreach (var kv in colVotes)
            {
                string gridHdr = kv.Key;
                string winnerLetter = kv.Value
                    .OrderByDescending(v => v.Value)
                    .First().Key;

                int colIndex = ExcelValueExtractor.ColIndexToLetter(1) == "A"
                    ? LetterToColIndex(winnerLetter)
                    : 1;   // fallback

                colXPositions.TryGetValue(gridHdr, out double screenX);

                mapping.Columns[gridHdr] = new GridColumnMapping
                {
                    GridColumnHeader = gridHdr,
                    ExcelColIndex = colIndex,
                    ExcelColLetter = winnerLetter,
                    GridColumnX = screenX,
                    CellType = InferCellType(gridHdr)
                };
            }

            return mapping;
        }

        // ── private: utility helpers ──────────────────────────────────────────

        private static Dictionary<string, List<ExcelCellValue>> BuildExcelLookup(
            List<ExcelCellValue> excelValues)
        {
            var lookup = new Dictionary<string, List<ExcelCellValue>>(StringComparer.Ordinal);
            foreach (var cell in excelValues)
            {
                string key = cell.NormalizedValue ?? string.Empty;
                if (string.IsNullOrEmpty(key)) continue;
                if (!lookup.TryGetValue(key, out var list))
                    lookup[key] = list = new List<ExcelCellValue>();
                list.Add(cell);
            }
            return lookup;
        }

        /// <summary>
        /// Builds the set of control names that belong to a grid (so they are
        /// excluded from header-field matching).
        /// </summary>
        private static HashSet<string> BuildGridControlNameSet(VisionSnapshot snapshot)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var tab in snapshot.Tabs)
            {
                if (tab.Grid == null) continue;
                foreach (var row in tab.Grid.Rows)
                    foreach (var cell in row.Cells)
                        if (!string.IsNullOrEmpty(cell.Value))
                            set.Add(cell.Value);
            }
            return set;
        }

        private static string InferInputMethod(string value)
        {
            if (string.IsNullOrEmpty(value)) return "SendKeys";

            // Multi-line or long values go via clipboard.
            if (value.Contains('\n') || value.Length > 50) return "Clipboard";

            // Circled numbers or short coded values may be dropdowns.
            bool hasCircled = value.Any(c => c >= '\u2460' && c <= '\u2473');
            if (hasCircled) return "Dropdown";

            return "SendKeys";
        }

        private static string InferCellType(string columnHeader)
        {
            if (string.IsNullOrEmpty(columnHeader)) return "text";

            string h = columnHeader.ToLowerInvariant();

            if (h.Contains("フラグ") || h.Contains("flag") ||
                h.Contains("有無") || h.Contains("可否"))
                return "checkbox";

            if (h.Contains("種類") || h.Contains("区分") ||
                h.Contains("type") || h.Contains("分類"))
                return "dropdown";

            if (h.Contains("数") || h.Contains("件数") ||
                h.Contains("金額") || h.Contains("桁"))
                return "numeric";

            return "text";
        }

        private static int LetterToColIndex(string letter)
        {
            int result = 0;
            foreach (char c in letter.ToUpperInvariant())
                result = result * 26 + (c - 'A' + 1);
            return result;
        }

        // Escapes a value for use inside a Markdown table cell.
        private static string Md(string s) =>
            (s ?? string.Empty).Replace("|", "\\|").Replace("\n", " ");
    }
}
