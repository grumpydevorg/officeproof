using System.Diagnostics;
using System.Globalization;
using System.Text;

namespace OfficeProof.Verify;

/// <summary>Changed pixels' bounding box, inclusive, in pixels at <see cref="PageDiff.Dpi"/> from the page's top-left corner.</summary>
public readonly record struct PixelBox(int Left, int Top, int Right, int Bottom);

/// <summary>One page of two PDFs compared as grey-level images.</summary>
/// <param name="Page">1-based page number.</param>
/// <param name="ChangedPixels">
/// Pixels whose grey levels differ by more than <see cref="PageDiff.PixelTolerance"/>. Where the
/// pages differ in size, every pixel outside the common area counts; a page present in one PDF
/// only counts every pixel.
/// </param>
/// <param name="ChangedFraction"><paramref name="ChangedPixels"/> over the pixels of the larger page.</param>
/// <param name="Box">Bounding box of the changed pixels; null when there are none.</param>
public sealed record PageDifference(int Page, long ChangedPixels, double ChangedFraction, PixelBox? Box)
{
    /// <summary>True when any pixel changed.</summary>
    public bool Differs => ChangedPixels > 0;
}

/// <summary>Two PDFs compared page by page.</summary>
/// <param name="PagesA">Pages in the first PDF.</param>
/// <param name="PagesB">Pages in the second PDF.</param>
/// <param name="Pages">One entry per page number up to the longer PDF, in order.</param>
public sealed record PageDiffResult(int PagesA, int PagesB, IReadOnlyList<PageDifference> Pages)
{
    /// <summary>Same page count and no changed pixel on any page.</summary>
    public bool Identical => PagesA == PagesB && Pages.All(p => !p.Differs);

    /// <summary>The pages with changed pixels.</summary>
    public IEnumerable<PageDifference> Differing => Pages.Where(p => p.Differs);
}

/// <summary>
/// Renders two PDFs with poppler's <c>pdftoppm</c> (on PATH) to 8-bit grey images and compares
/// them pixel by pixel. No image library: the PGM output is parsed here (the repository dropped
/// ImageSharp for a vulnerability).
/// </summary>
/// <remarks>
/// Noise threshold, measured 2026-10-15 (Word 16.113, poppler 26.09): Word's PDF of the same
/// document exported twice renders to identical page images (the TypedNumberReport fixture; the 42-page,
/// 19 MB client report; the report against Word's resave of it), so any changed pixel is a real
/// difference and <see cref="PixelTolerance"/> is 0. The comparison is that sensitive: a comment
/// anchor that moved the rest of one justified line 0.18 pt changed 2401 pixels in a 420 x 16
/// pixel box.
/// </remarks>
public static class PageDiff
{
    /// <summary>Rendering resolution. 100 dpi draws a 12 pt glyph about 17 pixels high and an A4 page in 1 MB.</summary>
    public const int Dpi = 100;

    /// <summary>Largest grey-level difference (0-255) still counted as unchanged.</summary>
    public const int PixelTolerance = 0;

    /// <summary>
    /// Compares <paramref name="pdfA"/> with <paramref name="pdfB"/>. With
    /// <paramref name="keepDir"/>, each differing page is written there as
    /// <c>page-NNN-a.pgm</c>, <c>page-NNN-b.pgm</c> and <c>page-NNN-diff.pgm</c> (changed pixels
    /// black over a faint copy of page A); PGM opens in Preview.
    /// </summary>
    public static PageDiffResult Compare(string pdfA, string pdfB, string? keepDir = null)
    {
        var work = Path.Combine(Path.GetTempPath(), "officeproof-pagediff-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(work);
        try
        {
            var renderA = Task.Run(() => Render(pdfA, Path.Combine(work, "a")));
            var renderB = Task.Run(() => Render(pdfB, Path.Combine(work, "b")));
            var pagesA = renderA.GetAwaiter().GetResult();
            var pagesB = renderB.GetAwaiter().GetResult();
            if (keepDir is not null) Directory.CreateDirectory(keepDir);

            var result = new List<PageDifference>();
            for (var i = 0; i < Math.Max(pagesA.Count, pagesB.Count); i++)
            {
                var a = i < pagesA.Count ? ReadPgm(pagesA[i]) : null;
                var b = i < pagesB.Count ? ReadPgm(pagesB[i]) : null;
                var (difference, diffImage) = ComparePage(i + 1, a, b);
                result.Add(difference);
                if (keepDir is null || !difference.Differs) continue;
                var stem = Path.Combine(keepDir, $"page-{i + 1:000}");
                if (a is not null) File.Copy(pagesA[i], stem + "-a.pgm", overwrite: true);
                if (b is not null) File.Copy(pagesB[i], stem + "-b.pgm", overwrite: true);
                if (diffImage is not null) WritePgm(stem + "-diff.pgm", diffImage);
            }
            return new PageDiffResult(pagesA.Count, pagesB.Count, result);
        }
        finally
        {
            try { Directory.Delete(work, recursive: true); }
            catch (IOException) { }
        }
    }

    /// <summary>An 8-bit grey image, rows top to bottom.</summary>
    internal sealed record GrayImage(int Width, int Height, byte[] Pixels);

    // The page's difference and, when both pages exist, a diff image the size of the larger one.
    internal static (PageDifference Difference, GrayImage? Diff) ComparePage(int page, GrayImage? a, GrayImage? b)
    {
        if (a is null || b is null)
        {
            var only = a ?? b ?? throw new ArgumentException("at least one page is required");
            var all = (long)only.Width * only.Height;
            return (new PageDifference(page, all, 1, new PixelBox(0, 0, only.Width - 1, only.Height - 1)), null);
        }
        var width = Math.Max(a.Width, b.Width);
        var height = Math.Max(a.Height, b.Height);
        var diff = new byte[width * height];
        long changed = 0;
        int left = int.MaxValue, top = int.MaxValue, right = -1, bottom = -1;
        for (var y = 0; y < height; y++)
        {
            for (var x = 0; x < width; x++)
            {
                var inA = x < a.Width && y < a.Height;
                var inB = x < b.Width && y < b.Height;
                var pa = inA ? a.Pixels[y * a.Width + x] : (byte)255;
                var pb = inB ? b.Pixels[y * b.Width + x] : (byte)255;
                var isChanged = !inA || !inB || Math.Abs(pa - pb) > PixelTolerance;
                // Faint copy of page A (a quarter of its darkness), changed pixels black.
                diff[y * width + x] = isChanged ? (byte)0 : (byte)(255 - (255 - pa) / 4);
                if (!isChanged) continue;
                changed++;
                if (x < left) left = x;
                if (x > right) right = x;
                if (y < top) top = y;
                if (y > bottom) bottom = y;
            }
        }
        var box = changed == 0 ? (PixelBox?)null : new PixelBox(left, top, right, bottom);
        return (new PageDifference(page, changed, (double)changed / ((long)width * height), box), new GrayImage(width, height, diff));
    }

    // Page images in page order. pdftoppm names them <prefix>-<n>.pgm, zero-padded to the
    // width of the page count.
    private static List<string> Render(string pdf, string prefix)
    {
        if (!File.Exists(pdf)) throw new FileNotFoundException("PDF not found", pdf);
        var exe = FindOnPath("pdftoppm") ?? throw new InvalidOperationException("pdftoppm (poppler) is not on PATH; it renders the PDF pages for comparison");
        var psi = new ProcessStartInfo(exe, ["-gray", "-r", Dpi.ToString(CultureInfo.InvariantCulture), pdf, prefix])
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
        };
        using var process = Process.Start(psi) ?? throw new InvalidOperationException("could not start pdftoppm");
        var stderr = process.StandardError.ReadToEndAsync();
        process.StandardOutput.ReadToEnd();
        process.WaitForExit();
        if (process.ExitCode != 0)
            throw new InvalidDataException($"pdftoppm failed on {pdf} (exit {process.ExitCode}): {stderr.Result.Trim()}");
        var name = Path.GetFileName(prefix) + "-";
        return [.. Directory.GetFiles(Path.GetDirectoryName(prefix)!, name + "*.pgm")
            .Select(f => (Path: f, Page: int.Parse(Path.GetFileNameWithoutExtension(f)[name.Length..], CultureInfo.InvariantCulture)))
            .OrderBy(p => p.Page)
            .Select(p => p.Path)];
    }

    private static string? FindOnPath(string name) =>
        (Environment.GetEnvironmentVariable("PATH") ?? "")
            .Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries)
            .Select(dir => Path.Combine(dir, name))
            .FirstOrDefault(File.Exists);

    /// <summary>Reads a binary 8-bit PGM (P5), the format pdftoppm writes with <c>-gray</c>.</summary>
    internal static GrayImage ReadPgm(string path)
    {
        var bytes = File.ReadAllBytes(path);
        var pos = 0;
        if (Token() != "P5") throw new InvalidDataException($"{path}: not a binary PGM (P5)");
        var width = Number();
        var height = Number();
        var maxValue = Number();
        if (maxValue is < 1 or > 255) throw new InvalidDataException($"{path}: only 8-bit PGM is supported (maxval {maxValue})");
        pos++; // the single whitespace byte before the raster
        var size = width * height;
        if (bytes.Length - pos < size) throw new InvalidDataException($"{path}: raster truncated ({bytes.Length - pos} of {size} bytes)");
        return new GrayImage(width, height, bytes.AsSpan(pos, size).ToArray());

        string Token()
        {
            while (pos < bytes.Length)
            {
                if (bytes[pos] == '#') { while (pos < bytes.Length && bytes[pos] != '\n') pos++; }
                else if (char.IsWhiteSpace((char)bytes[pos])) pos++;
                else break;
            }
            var start = pos;
            while (pos < bytes.Length && !char.IsWhiteSpace((char)bytes[pos])) pos++;
            return Encoding.ASCII.GetString(bytes, start, pos - start);
        }

        int Number() => int.TryParse(Token(), NumberStyles.None, CultureInfo.InvariantCulture, out var n) && n > 0
            ? n
            : throw new InvalidDataException($"{path}: malformed PGM header");
    }

    internal static void WritePgm(string path, GrayImage image)
    {
        using var file = File.Create(path);
        file.Write(Encoding.ASCII.GetBytes(string.Create(CultureInfo.InvariantCulture, $"P5\n{image.Width} {image.Height}\n255\n")));
        file.Write(image.Pixels);
    }
}
