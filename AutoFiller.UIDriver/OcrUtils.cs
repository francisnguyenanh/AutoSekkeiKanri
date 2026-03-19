using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AutoFiller.UIDriver
{
    /// <summary>
    /// Shared utilities for OCR word analysis and value normalisation.
    /// Centralises logic that would otherwise be duplicated across
    /// <see cref="VisionInteractor"/>, <see cref="TabNavigator"/>, and
    /// <see cref="AppSnapshotDumper"/>.
    /// </summary>
    public static class OcrUtils
    {
        /// <summary>
        /// Groups OCR words into horizontal bands by Y-centre proximity.
        /// Words whose Y-centres are within <paramref name="tolerance"/> px are
        /// placed in the same band.  Returns bands sorted top-to-bottom.
        /// </summary>
        public static List<List<OcrWord>> GroupByY(IEnumerable<OcrWord> words, int tolerance = 8)
        {
            var bands = new List<List<OcrWord>>();
            foreach (var word in words.OrderBy(w => w.BoundingBox.Top))
            {
                int cy = word.BoundingBox.Top + word.BoundingBox.Height / 2;
                var band = bands.FirstOrDefault(b =>
                    Math.Abs(cy - (b[0].BoundingBox.Top + b[0].BoundingBox.Height / 2)) <= tolerance);
                if (band != null) band.Add(word);
                else              bands.Add(new List<OcrWord> { word });
            }
            return bands;
        }

        /// <summary>
        /// Finds the band most likely to be a column header row:
        /// highest word count with consistent horizontal gaps.
        /// Returns <c>null</c> when <paramref name="bands"/> contains no band
        /// with at least two words.
        /// </summary>
        public static List<OcrWord> FindHeaderBand(List<List<OcrWord>> bands)
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
        /// Normalises a string value for OCR comparison:
        /// trims whitespace, removes full-width spaces (U+3000), and converts
        /// full-width digits ０–９ (U+FF10–U+FF19) to half-width ASCII digits.
        /// </summary>
        public static string NormalizeValue(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            value = value.Trim().Replace("\u3000", " ").Trim();
            // Full-width digit → half-width: ０-９ (U+FF10–U+FF19) → 0-9
            var sb = new StringBuilder(value.Length);
            foreach (char c in value)
                sb.Append(c >= '\uFF10' && c <= '\uFF19' ? (char)(c - '\uFF10' + '0') : c);
            return sb.ToString();
        }
    }
}
