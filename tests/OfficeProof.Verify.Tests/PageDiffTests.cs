using System.Globalization;
using System.Text;

namespace OfficeProof.Verify.Tests;

public sealed class PageDiffTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "officeproof-pagediff-tests", Guid.NewGuid().ToString("N"));

    public PageDiffTests() => Directory.CreateDirectory(_dir);

    public void Dispose()
    {
        try { Directory.Delete(_dir, recursive: true); }
        catch (IOException) { }
    }

    private static PageDiff.GrayImage Blank(int width, int height) => new(width, height, Enumerable.Repeat((byte)255, width * height).ToArray());

    [Fact]
    public void Changed_pixels_are_counted_and_bounded()
    {
        var a = Blank(10, 8);
        var b = Blank(10, 8);
        b.Pixels[2 * 10 + 3] = 0;
        b.Pixels[5 * 10 + 7] = 254;

        var (difference, diff) = PageDiff.ComparePage(1, a, b);

        Assert.Equal(2, difference.ChangedPixels);
        Assert.Equal(new PixelBox(3, 2, 7, 5), difference.Box);
        Assert.Equal(2.0 / 80, difference.ChangedFraction, 9);
        Assert.Equal(0, diff!.Pixels[2 * 10 + 3]);
        Assert.Equal(255, diff.Pixels[0]);
    }

    [Fact]
    public void Identical_pages_have_no_box()
    {
        var (difference, _) = PageDiff.ComparePage(3, Blank(4, 4), Blank(4, 4));

        Assert.False(difference.Differs);
        Assert.Null(difference.Box);
    }

    [Fact]
    public void A_page_in_one_pdf_only_is_wholly_changed()
    {
        var (difference, diff) = PageDiff.ComparePage(2, Blank(5, 4), null);

        Assert.Equal(20, difference.ChangedPixels);
        Assert.Equal(1, difference.ChangedFraction);
        Assert.Equal(new PixelBox(0, 0, 4, 3), difference.Box);
        Assert.Null(diff);
    }

    [Fact]
    public void Pages_of_different_size_count_the_uncovered_area()
    {
        var (difference, _) = PageDiff.ComparePage(1, Blank(4, 4), Blank(4, 5));

        Assert.Equal(4, difference.ChangedPixels);
        Assert.Equal(new PixelBox(0, 4, 3, 4), difference.Box);
    }

    [Fact]
    public void Pgm_round_trips_and_skips_header_comments()
    {
        var image = new PageDiff.GrayImage(3, 2, [0, 10, 20, 30, 40, 255]);
        var path = Path.Combine(_dir, "x.pgm");
        PageDiff.WritePgm(path, image);
        Assert.Equal(image.Pixels, PageDiff.ReadPgm(path).Pixels);

        var commented = Path.Combine(_dir, "c.pgm");
        File.WriteAllBytes(commented, [.. Encoding.ASCII.GetBytes("P5\n# made by hand\n3 2\n255\n"), .. image.Pixels]);
        var read = PageDiff.ReadPgm(commented);
        Assert.Equal((3, 2), (read.Width, read.Height));
        Assert.Equal(image.Pixels, read.Pixels);
    }

    [Fact]
    public void Non_pgm_input_is_refused()
    {
        var path = Path.Combine(_dir, "p6.ppm");
        File.WriteAllBytes(path, Encoding.ASCII.GetBytes("P6\n1 1\n255\n\0\0\0"));

        Assert.Throws<InvalidDataException>(() => PageDiff.ReadPgm(path));
    }

    internal static void RequirePdftoppm() =>
        Assert.SkipWhen(PageDiff.FindOnPath("pdftoppm") is null, "pdftoppm (poppler) is not on PATH");

    [Fact]
    public void A_command_is_found_on_path_with_a_pathext_extension_where_one_is_given()
    {
        var first = Directory.CreateDirectory(Path.Combine(_dir, "first")).FullName;
        var second = Directory.CreateDirectory(Path.Combine(_dir, "second")).FullName;
        var bare = Path.Combine(first, "tool");
        var exe = Path.Combine(second, "tool.EXE");
        File.WriteAllBytes(bare, []);
        File.WriteAllBytes(exe, []);
        var path = string.Join(Path.PathSeparator, $"\"{first}\"", second);

        Assert.Equal(exe, PageDiff.FindOnPath("tool", path, ".COM;.EXE"));
        Assert.Equal(bare, PageDiff.FindOnPath("tool", path, pathExt: null));
        Assert.Equal(exe, PageDiff.FindOnPath("tool.EXE", path, ".COM;.EXE"));
        Assert.Null(PageDiff.FindOnPath("other", path, ".COM;.EXE"));
    }

    [Fact]
    public void Same_pdf_content_renders_identically()
    {
        RequirePdftoppm();
        var a = MinimalPdf(Path.Combine(_dir, "a.pdf"), "Alpha beta gamma", "Second page");
        var b = MinimalPdf(Path.Combine(_dir, "b.pdf"), "Alpha beta gamma", "Second page");

        var result = PageDiff.Compare(a, b);

        Assert.True(result.Identical);
        Assert.Equal((2, 2), (result.PagesA, result.PagesB));
    }

    [Fact]
    public void A_changed_word_is_found_on_its_page_and_kept()
    {
        RequirePdftoppm();
        var a = MinimalPdf(Path.Combine(_dir, "a.pdf"), "Same first page", "Alpha beta gamma");
        var b = MinimalPdf(Path.Combine(_dir, "b.pdf"), "Same first page", "Alpha delta gamma");
        var keep = Path.Combine(_dir, "keep");

        var result = PageDiff.Compare(a, b, keep);

        var page = Assert.Single(result.Differing);
        Assert.Equal(2, page.Page);
        // The text starts 20 pt from the left edge and sits on a baseline 40 pt above the bottom
        // of a 100 pt page: at 100 dpi the change lies right of x = 28 and around y = 83.
        Assert.InRange(page.Box!.Value.Left, 28, 200);
        Assert.InRange(page.Box.Value.Top, 40, 83);
        Assert.InRange(page.Box.Value.Bottom, 60, 100);
        Assert.Equal(["page-002-a.pgm", "page-002-b.pgm", "page-002-diff.pgm"],
            Directory.GetFiles(keep).Select(Path.GetFileName).Order(StringComparer.Ordinal));
    }

    [Fact]
    public void An_extra_page_is_reported()
    {
        RequirePdftoppm();
        var a = MinimalPdf(Path.Combine(_dir, "a.pdf"), "One");
        var b = MinimalPdf(Path.Combine(_dir, "b.pdf"), "One", "Two");

        var result = PageDiff.Compare(a, b);

        Assert.Equal((1, 2), (result.PagesA, result.PagesB));
        Assert.False(result.Identical);
        Assert.Equal(2, Assert.Single(result.Differing).Page);
    }

    /// <summary>A PDF with one 200 x 100 pt page per string, each line in 24 pt Helvetica.</summary>
    internal static string MinimalPdf(string path, params string[] pages)
    {
        var objects = new List<string>
        {
            "<< /Type /Catalog /Pages 2 0 R >>",
            $"<< /Type /Pages /Kids [{string.Join(" ", pages.Select((_, i) => $"{4 + 2 * i} 0 R"))}] /Count {pages.Length} >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
        };
        foreach (var text in pages)
        {
            var content = $"BT /F1 24 Tf 20 40 Td ({text}) Tj ET";
            objects.Add($"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] /Contents {objects.Count + 2} 0 R /Resources << /Font << /F1 3 0 R >> >> >>");
            objects.Add($"<< /Length {content.Length} >>\nstream\n{content}\nendstream");
        }
        var pdf = new StringBuilder("%PDF-1.4\n");
        var offsets = new List<int>();
        for (var i = 0; i < objects.Count; i++)
        {
            offsets.Add(pdf.Length);
            pdf.Append(CultureInfo.InvariantCulture, $"{i + 1} 0 obj\n{objects[i]}\nendobj\n");
        }
        var xref = pdf.Length;
        pdf.Append(CultureInfo.InvariantCulture, $"xref\n0 {objects.Count + 1}\n0000000000 65535 f \n");
        foreach (var offset in offsets) pdf.Append(CultureInfo.InvariantCulture, $"{offset:0000000000} 00000 n \n");
        pdf.Append(CultureInfo.InvariantCulture, $"trailer\n<< /Size {objects.Count + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n");
        File.WriteAllText(path, pdf.ToString(), Encoding.ASCII);
        return path;
    }
}
