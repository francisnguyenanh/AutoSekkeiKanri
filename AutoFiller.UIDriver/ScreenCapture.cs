using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;

namespace AutoFiller.UIDriver
{
    /// <summary>
    /// Win32-backed window screenshot helper.
    /// Captures the target app window into a <see cref="Bitmap"/> using
    /// <c>PrintWindow(PW_RENDERFULLCONTENT)</c> so off-screen and partially
    /// obscured windows render correctly.
    /// </summary>
    public class ScreenCapture
    {
        // ── Win32 imports ─────────────────────────────────────────────────

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
        private delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hwnd, System.Text.StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hwnd);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hwnd);

        [DllImport("user32.dll")]
        private static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, uint nFlags);

        [DllImport("gdi32.dll")]
        private static extern bool BitBlt(IntPtr hdc, int x, int y, int cx, int cy,
                                          IntPtr hdcSrc, int x1, int y1, uint rop);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        // PrintWindow flag — renders the full window content including off-screen areas.
        private const uint PW_RENDERFULLCONTENT = 2;

        // ── public API ────────────────────────────────────────────────────

        /// <summary>
        /// Searches all top-level windows for one whose title contains
        /// <paramref name="titleContains"/> (case-insensitive).
        /// Returns <see cref="IntPtr.Zero"/> when no match is found.
        /// </summary>
        public IntPtr FindWindowByTitle(string titleContains)
        {
            if (string.IsNullOrEmpty(titleContains))
                throw new ArgumentException("titleContains must not be empty.", nameof(titleContains));

            IntPtr found = IntPtr.Zero;

            EnumWindows((hwnd, _) =>
            {
                if (!IsWindowVisible(hwnd)) return true; // continue

                var sb = new System.Text.StringBuilder(512);
                GetWindowText(hwnd, sb, sb.Capacity);
                string title = sb.ToString();

                if (title.IndexOf(titleContains, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    found = hwnd;
                    return false; // stop enumeration
                }
                return true;
            }, IntPtr.Zero);

            return found;
        }

        /// <summary>
        /// Captures the full client area of the window identified by
        /// <paramref name="hwnd"/> into a <see cref="Bitmap"/>.
        /// Uses <c>PrintWindow(PW_RENDERFULLCONTENT)</c> to render correctly even
        /// when the window is partially obscured or minimised.
        /// The returned bitmap is in screen coordinates — the window's top-left
        /// corner is (0, 0) within the bitmap.
        /// </summary>
        public Bitmap CaptureWindow(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero)
                throw new ArgumentException("hwnd must not be IntPtr.Zero.", nameof(hwnd));

            GetWindowRect(hwnd, out RECT rect);
            int width  = rect.Right  - rect.Left;
            int height = rect.Bottom - rect.Top;

            if (width <= 0 || height <= 0)
                throw new InvalidOperationException(
                    $"Window has invalid dimensions ({width}×{height}).");

            var bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb);
            using var g = Graphics.FromImage(bmp);
            IntPtr hdc = g.GetHdc();
            try
            {
                bool ok = PrintWindow(hwnd, hdc, PW_RENDERFULLCONTENT);
                if (!ok)
                {
                    // Fallback: BitBlt from screen (requires window to be visible).
                    using var screenGraphics = Graphics.FromHwnd(IntPtr.Zero);
                    IntPtr screenDc = screenGraphics.GetHdc();
                    try
                    {
                        const uint SRCCOPY = 0x00CC0020;
                        BitBlt(hdc, 0, 0, width, height, screenDc, rect.Left, rect.Top, SRCCOPY);
                    }
                    finally
                    {
                        screenGraphics.ReleaseHdc(screenDc);
                    }
                }
            }
            finally
            {
                g.ReleaseHdc(hdc);
            }

            return bmp;
        }

        /// <summary>
        /// Returns the window's bounding rectangle in screen coordinates.
        /// </summary>
        public Rectangle GetWindowBounds(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero)
                throw new ArgumentException("hwnd must not be IntPtr.Zero.", nameof(hwnd));

            GetWindowRect(hwnd, out RECT rect);
            return new Rectangle(rect.Left, rect.Top,
                                  rect.Right - rect.Left,
                                  rect.Bottom - rect.Top);
        }

        /// <summary>
        /// Captures a sub-region of the window.
        /// <paramref name="region"/> is in window-relative coordinates
        /// (0,0 = top-left of the window).
        /// </summary>
        public Bitmap CaptureRegion(IntPtr hwnd, Rectangle region)
        {
            using Bitmap full = CaptureWindow(hwnd);

            // Clamp to window bounds.
            int x = Math.Max(0, region.X);
            int y = Math.Max(0, region.Y);
            int w = Math.Min(region.Width,  full.Width  - x);
            int h = Math.Min(region.Height, full.Height - y);

            if (w <= 0 || h <= 0)
                throw new ArgumentOutOfRangeException(nameof(region),
                    "Region falls entirely outside the window bounds.");

            return full.Clone(new Rectangle(x, y, w, h), full.PixelFormat);
        }

        /// <summary>
        /// Saves <paramref name="bmp"/> to <paramref name="path"/>.
        /// The file format is inferred from the extension (.png, .jpg, .bmp).
        /// Defaults to PNG for unknown extensions.
        /// </summary>
        public void SaveScreenshot(Bitmap bmp, string path)
        {
            if (bmp == null)  throw new ArgumentNullException(nameof(bmp));
            if (string.IsNullOrEmpty(path))
                throw new ArgumentException("path must not be empty.", nameof(path));

            Directory.CreateDirectory(
                Path.GetDirectoryName(Path.GetFullPath(path))!);

            ImageFormat fmt = Path.GetExtension(path).ToLowerInvariant() switch
            {
                ".jpg" or ".jpeg" => ImageFormat.Jpeg,
                ".bmp"            => ImageFormat.Bmp,
                ".gif"            => ImageFormat.Gif,
                _                 => ImageFormat.Png   // default
            };

            bmp.Save(path, fmt);
        }
    }
}
