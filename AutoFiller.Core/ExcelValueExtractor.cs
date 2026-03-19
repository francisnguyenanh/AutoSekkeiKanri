// Requires NuGet package: EPPlus (>= 5.x)
// EPPlus 5+ licence notice: set ExcelPackage.LicenseContext before use.
//   Non-commercial: ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
//   Commercial:     ExcelPackage.LicenseContext = LicenseContext.Commercial;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using AutoFiller.UIDriver;
using OfficeOpenXml;

namespace AutoFiller.Core
{
    // ─────────────────────────────────────────────
    // Data model
    // ─────────────────────────────────────────────

    public class ExcelCellValue
    {
        /// <summary>Sheet (tab) name.</summary>
        public string SheetName { get; set; }

        /// <summary>1-based row index.</summary>
        public int Row { get; set; }

        /// <summary>1-based column index.</summary>
        public int Col { get; set; }

        /// <summary>Column letter(s): "A", "B", "AD" etc.</summary>
        public string ColLetter { get; set; }

        /// <summary>Cell address: "B7", "F12" etc.</summary>
        public string CellAddress { get; set; }

        /// <summary>Original value read from the cell, as a string.</summary>
        public string RawValue { get; set; }

        /// <summary>Trimmed and normalised value, ready for matching.</summary>
        public string NormalizedValue { get; set; }

        /// <summary>
        /// If the cell is the top-left of a merged range, contains the range
        /// address (e.g. "B7:D9"). Null for unmerged cells.
        /// </summary>
        public string MergedRangeAddress { get; set; }

        /// <summary>
        /// True when this entry was re-emitted for a non-top-left cell inside a
        /// merge (the top-left value is propagated). False for top-left cells and
        /// unmerged cells.
        /// </summary>
        public bool IsFromMerge { get; set; }

        /// <summary>"HEADER" for rows &lt; 47; "ITEM_TABLE" for rows &ge; 47.</summary>
        public string SectionHint { get; set; }
    }

    public class ItemRowValues
    {
        /// <summary>1-based Excel row number of the data row.</summary>
        public int SourceRow { get; set; }

        /// <summary>
        /// Value from the 番号 / 項番 column (①②③ …).
        /// Empty string when the column is absent.
        /// </summary>
        public string 番号 { get; set; }

        /// <summary>
        /// All non-empty cell values for this row, keyed by column header text
        /// (taken from row 47).
        /// </summary>
        public Dictionary<string, string> Values { get; set; } = new Dictionary<string, string>();
    }

    // ─────────────────────────────────────────────
    // Extractor
    // ─────────────────────────────────────────────

    public class ExcelValueExtractor
    {
        // Row boundary: rows < 47 are HEADER; rows >= 47 are ITEM_TABLE.
        private const int ItemTableStartRow = 47;

        // Column header row for the item table (first ITEM_TABLE row).
        private const int ColumnHeaderRow = 47;

        // Data rows start one row after the column header row.
        private const int DataStartRow = 48;

        // Template placeholder values that carry no real data.
        private static readonly HashSet<string> SkipValues = new HashSet<string>(
            StringComparer.Ordinal)
        {
            "－", "-", "※"
        };

        private static readonly string[] SkipPrefixes =
        {
            "＜変更前＞", "＜変更後＞"
        };

        // ── public API ────────────────────────────────────────────────────────

        /// <summary>
        /// Extracts ALL non-empty, non-template cells from <paramref name="sheetName"/>
        /// in the workbook at <paramref name="filePath"/>.
        /// For merged ranges only the top-left cell is emitted.
        /// </summary>
        public List<ExcelCellValue> ExtractAll(string filePath, string sheetName)
        {
            ValidateFile(filePath);

            using var package = OpenPackage(filePath);
            var sheet = GetSheet(package, sheetName);
            return ExtractFromSheet(sheet, sheetName, minRow: 1, maxRow: int.MaxValue);
        }

        /// <summary>
        /// Extracts cells from rows 1–47 (the fixed-field header section).
        /// </summary>
        public List<ExcelCellValue> ExtractHeaderValues(string filePath, string sheetName)
        {
            ValidateFile(filePath);

            using var package = OpenPackage(filePath);
            var sheet = GetSheet(package, sheetName);
            return ExtractFromSheet(sheet, sheetName, minRow: 1, maxRow: 47);
        }

        /// <summary>
        /// Extracts rows 48+ and groups them by Excel row.
        /// Column headers are read from row 47.
        /// Each returned <see cref="ItemRowValues"/> maps column-header → cell value.
        /// </summary>
        public List<ItemRowValues> ExtractItemRows(string filePath, string sheetName)
        {
            ValidateFile(filePath);

            using var package = OpenPackage(filePath);
            var sheet = GetSheet(package, sheetName);

            // ── Column headers from row 47 ─────────────────────────────────
            // key: 1-based column index  |  value: header text
            var columnHeaders = new Dictionary<int, string>();

            if (sheet.Dimension != null)
            {
                int lastCol = sheet.Dimension.End.Column;
                for (int col = 1; col <= lastCol; col++)
                {
                    string hdr = CellText(sheet, ColumnHeaderRow, col);
                    if (!string.IsNullOrWhiteSpace(hdr))
                        columnHeaders[col] = hdr.Trim();
                }
            }

            // Identify the 番号 column (circled-number / sequence column).
            int bango番号Col = ResolveBangoColumn(columnHeaders);

            // ── Build merge map for fast look-up ───────────────────────────
            var (topLeftMap, _) = BuildMergeMap(sheet);

            // ── Collect data rows ──────────────────────────────────────────
            var result = new List<ItemRowValues>();

            if (sheet.Dimension == null) return result;

            int lastRow = sheet.Dimension.End.Row;
            int lastDataCol = sheet.Dimension.End.Column;

            for (int row = DataStartRow; row <= lastRow; row++)
            {
                var rowEntry = new ItemRowValues { SourceRow = row };
                bool hasAnyValue = false;

                for (int col = 1; col <= lastDataCol; col++)
                {
                    // If this cell is inside a merge but is NOT the top-left, skip.
                    // The top-left already represents the merged value.
                    if (IsNonTopLeftMergeCell(sheet, row, col, topLeftMap))
                        continue;

                    string raw = CellText(sheet, row, col);
                    if (string.IsNullOrWhiteSpace(raw) || ShouldSkip(raw))
                        continue;

                    hasAnyValue = true;
                    string colHdr = columnHeaders.TryGetValue(col, out string h) ? h : ColIndexToLetter(col);

                    if (col == bango番号Col)
                        rowEntry.番号 = raw.Trim();
                    else
                        rowEntry.Values[colHdr] = raw.Trim();
                }

                if (hasAnyValue)
                    result.Add(rowEntry);
            }

            return result;
        }

        /// <summary>
        /// Normalises a raw cell value for fuzzy / case-insensitive matching:
        /// trims whitespace + full-width spaces, converts full-width digits and
        /// ASCII letters to half-width, lowercases, strips trailing ".0".
        /// </summary>
        public string Normalize(string value)
        {
            if (value == null) return string.Empty;

            // 1. Trim + full-width spaces + full-width digits → shared utility.
            value = OcrUtils.NormalizeValue(value);

            // 2. Full-width Latin letters → half-width (Excel-specific).
            var sb = new StringBuilder(value.Length);
            foreach (char c in value)
            {
                if      (c >= '\uFF21' && c <= '\uFF3A') sb.Append((char)(c - '\uFF21' + 'A'));
                else if (c >= '\uFF41' && c <= '\uFF5A') sb.Append((char)(c - '\uFF41' + 'a'));
                else sb.Append(c);
            }
            value = sb.ToString();

            // 3. Lowercase.
            value = value.ToLowerInvariant();

            // 4. Strip trailing ".0" from numbers that Excel may have serialised
            //    as floats (e.g. "2" stored as 2.0 → "2.0" → "2").
            if (value.EndsWith(".0", StringComparison.Ordinal))
                value = value[..^2];

            return value;
        }

        // ── private helpers ───────────────────────────────────────────────────

        /// <summary>
        /// Core extraction loop. Handles merged cells, skip rules, section hints.
        /// Only the top-left cell of each merged range is emitted.
        /// </summary>
        private List<ExcelCellValue> ExtractFromSheet(
            ExcelWorksheet sheet, string sheetName, int minRow, int maxRow)
        {
            var result = new List<ExcelCellValue>();
            if (sheet.Dimension == null) return result;

            int lastRow = Math.Min(sheet.Dimension.End.Row, maxRow);
            int lastCol = sheet.Dimension.End.Column;

            var (topLeftMap, mergeRangeMap) = BuildMergeMap(sheet);

            for (int row = Math.Max(1, minRow); row <= lastRow; row++)
            {
                for (int col = 1; col <= lastCol; col++)
                {
                    // Skip non-top-left cells within a merge.
                    if (IsNonTopLeftMergeCell(sheet, row, col, topLeftMap))
                        continue;

                    string raw = CellText(sheet, row, col);
                    if (string.IsNullOrWhiteSpace(raw) || ShouldSkip(raw))
                        continue;

                    string colLetter = ColIndexToLetter(col);
                    string cellAddr = $"{colLetter}{row}";

                    // Find the merge range this cell belongs to (if any).
                    mergeRangeMap.TryGetValue((row, col), out string mergeAddr);

                    result.Add(new ExcelCellValue
                    {
                        SheetName = sheetName,
                        Row = row,
                        Col = col,
                        ColLetter = colLetter,
                        CellAddress = cellAddr,
                        RawValue = raw,
                        NormalizedValue = Normalize(raw),
                        MergedRangeAddress = mergeAddr,
                        IsFromMerge = false,
                        SectionHint = row < ItemTableStartRow ? "HEADER" : "ITEM_TABLE"
                    });
                }
            }

            return result;
        }

        /// <summary>
        /// Builds two maps from the worksheet's merged-cell list:
        /// <list type="bullet">
        ///   <item><c>topLeftMap</c>: (row, col) → (topRow, topCol) for every cell inside
        ///   a merge range (including the top-left itself).</item>
        ///   <item><c>mergeRangeMap</c>: (topRow, topCol) → range-address-string for the
        ///   top-left cell only.</item>
        /// </list>
        /// </summary>
        private static (
            Dictionary<(int, int), (int, int)> topLeftMap,
            Dictionary<(int, int), string> mergeRangeMap)
            BuildMergeMap(ExcelWorksheet sheet)
        {
            var topLeftMap = new Dictionary<(int, int), (int, int)>();
            var mergeRangeMap = new Dictionary<(int, int), string>();

            if (sheet.MergedCells == null) return (topLeftMap, mergeRangeMap);

            foreach (string rangeAddr in sheet.MergedCells)
            {
                if (string.IsNullOrEmpty(rangeAddr)) continue;

                var range = sheet.Cells[rangeAddr];
                int topRow = range.Start.Row;
                int topCol = range.Start.Column;

                mergeRangeMap[(topRow, topCol)] = rangeAddr;

                for (int r = range.Start.Row; r <= range.End.Row; r++)
                for (int c = range.Start.Column; c <= range.End.Column; c++)
                    topLeftMap[(r, c)] = (topRow, topCol);
            }

            return (topLeftMap, mergeRangeMap);
        }

        /// <summary>
        /// Returns true when (row, col) is part of a merge range but is NOT the
        /// top-left cell (and should therefore be skipped).
        /// </summary>
        private static bool IsNonTopLeftMergeCell(
            ExcelWorksheet sheet,
            int row, int col,
            Dictionary<(int, int), (int, int)> topLeftMap)
        {
            if (!topLeftMap.TryGetValue((row, col), out var tl)) return false;
            return tl.Item1 != row || tl.Item2 != col;
        }

        /// <summary>
        /// Returns the cell's text representation. For formula cells, returns
        /// the cached display value. For numeric cells, avoids ".0" noise by
        /// using the formatted text when available.
        /// </summary>
        private static string CellText(ExcelWorksheet sheet, int row, int col)
        {
            var cell = sheet.Cells[row, col];
            if (cell.Value == null) return string.Empty;

            // Prefer the formatted text string that Excel would display.
            string text = cell.Text;
            if (!string.IsNullOrEmpty(text)) return text;

            return cell.Value.ToString() ?? string.Empty;
        }

        /// <summary>
        /// Returns true for values that are Excel template placeholders rather
        /// than real data.
        /// </summary>
        private static bool ShouldSkip(string value)
        {
            string trimmed = value.Trim();
            if (SkipValues.Contains(trimmed)) return true;
            foreach (string prefix in SkipPrefixes)
                if (trimmed.StartsWith(prefix, StringComparison.Ordinal)) return true;
            return false;
        }

        /// <summary>
        /// Finds the 1-based column index whose header indicates a 番号 / 項番
        /// (circled-number / sequence) column. Returns 0 if not found.
        /// </summary>
        private static int ResolveBangoColumn(Dictionary<int, string> headers)
        {
            string[] candidates = { "番号", "項番", "No.", "No", "№" };
            foreach (var kv in headers)
                foreach (string candidate in candidates)
                    if (kv.Value.Contains(candidate, StringComparison.Ordinal))
                        return kv.Key;
            return 0;
        }

        // ── Excel address helpers ─────────────────────────────────────────────

        /// <summary>
        /// Converts a 1-based column index to an Excel column letter string
        /// (1 → "A", 26 → "Z", 27 → "AA", …).
        /// </summary>
        public static string ColIndexToLetter(int col)
        {
            if (col <= 0) throw new ArgumentOutOfRangeException(nameof(col));

            var sb = new StringBuilder();
            while (col > 0)
            {
                col--;
                sb.Insert(0, (char)('A' + col % 26));
                col /= 26;
            }
            return sb.ToString();
        }

        // ── file / package helpers ────────────────────────────────────────────

        private static void ValidateFile(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                throw new ArgumentException("filePath must not be empty.", nameof(filePath));
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Excel file not found.", filePath);
        }

        private static ExcelPackage OpenPackage(string filePath)
        {
            // EPPlus 5+ requires a licence context to be set before first use.
            // Default to non-commercial; consuming code can override before calling.
            if (ExcelPackage.LicenseContext == null)
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            return new ExcelPackage(new FileInfo(filePath));
        }

        private static ExcelWorksheet GetSheet(ExcelPackage package, string sheetName)
        {
            var ws = package.Workbook.Worksheets[sheetName];
            if (ws == null)
                throw new ArgumentException(
                    $"Sheet \"{sheetName}\" not found in workbook.", nameof(sheetName));
            return ws;
        }
    }
}
