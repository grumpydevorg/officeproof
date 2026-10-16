using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeProof.Core;

namespace OfficeProof.Templating;

/// <summary>Which <c>w:font/w:embed*</c> slot a face fills. Declaration order is the schema order.</summary>
public enum FontSlot { Regular, Bold, Italic, BoldItalic }

/// <summary>One font file and the slot it fills.</summary>
public sealed record FontFace(string FileName, FontSlot Slot);

/// <summary>Outcome of <see cref="FontEmbedder.Embed"/>.</summary>
/// <param name="Family">The font family name written to the font table.</param>
/// <param name="FontDirectories">Where the faces were looked for, in search order.</param>
/// <param name="Embedded">File names embedded.</param>
/// <param name="Missing">File names not found; the document still builds, rendering falls back on machines without the font.</param>
public sealed record FontEmbedResult(string Family, IReadOnlyList<string> FontDirectories, IReadOnlyList<string> Embedded, IReadOnlyList<string> Missing);

/// <summary>
/// Embeds a font family as obfuscated <c>.odttf</c> font parts so the document renders in it on
/// machines that do not have it installed. Word only uses them when settings carry
/// <c>w:embedTrueTypeFonts</c> (see <see cref="SettingsInjector"/>).
/// </summary>
public sealed class FontEmbedder(IReadOnlyList<string> fontDirectories, string family, IReadOnlyList<FontFace> faces)
{
    private const string ObfuscatedFont = "application/vnd.openxmlformats-officedocument.obfuscatedFont";
    private static readonly XNamespace W = Ns.W;
    private static readonly XNamespace R = Ns.R;
    private static readonly EnumerationOptions Subdirectories = new() { RecurseSubdirectories = true, IgnoreInaccessible = true };

    /// <summary>
    /// Where faces are looked for. A face is the first file with its name in the first directory
    /// that has one, at the top level or else in a subdirectory (Linux keeps system fonts in
    /// trees such as /usr/share/fonts/truetype/&lt;family&gt;/). Directories that do not exist are skipped.
    /// </summary>
    public IReadOnlyList<string> FontDirectories { get; } = [.. fontDirectories.Select(Path.GetFullPath)];
    public string Family { get; } = family;
    public IReadOnlyList<FontFace> Faces { get; } = [.. faces.OrderBy(f => f.Slot)];

    /// <summary>
    /// Where a template's faces are looked for by default: the user's font directory, then the
    /// machine's. macOS: <c>~/Library/Fonts</c>, <c>/Library/Fonts</c>. Windows:
    /// <c>%LOCALAPPDATA%\Microsoft\Windows\Fonts</c> (fonts installed for the current user),
    /// <c>%WINDIR%\Fonts</c>. Linux and other Unix: <c>$XDG_DATA_HOME/fonts</c> (default
    /// <c>~/.local/share/fonts</c>), <c>/usr/local/share/fonts</c>, <c>/usr/share/fonts</c>.
    /// </summary>
    public static IReadOnlyList<string> DefaultFontDirectories
    {
        get
        {
            if (OperatingSystem.IsWindows())
            {
                return [Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Windows", "Fonts"),
                    Environment.GetFolderPath(Environment.SpecialFolder.Fonts)];
            }
            var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            if (OperatingSystem.IsMacOS()) return [Path.Combine(home, "Library", "Fonts"), "/Library/Fonts"];
            var dataHome = Environment.GetEnvironmentVariable("XDG_DATA_HOME") is { } x && Path.IsPathFullyQualified(x)
                ? x
                : Path.Combine(home, ".local", "share");
            return [Path.Combine(dataHome, "fonts"), "/usr/local/share/fonts", "/usr/share/fonts"];
        }
    }

    /// <summary>
    /// Embeds every face found in <see cref="FontDirectories"/>. Missing faces are listed in the
    /// result, never thrown. Any existing entry for <see cref="Family"/> in the font table is
    /// replaced (its font parts deleted); other fonts in the table are kept.
    /// </summary>
    public FontEmbedResult Embed(MainDocumentPart main)
    {
        var table = main.FontTablePart;
        XDocument xml;
        if (table is null)
        {
            table = main.AddNewPart<FontTablePart>();
            xml = NewTable();
        }
        else
        {
            xml = TryLoad(table) ?? NewTable();
        }
        var root = xml.Root!;
        if (root.GetPrefixOfNamespace(R) is null)
            root.SetAttributeValue(XNamespace.Xmlns + "r", R.NamespaceName);

        // Re-embedding (retemplating a document officeproof already produced) must not leave the
        // previous run's font parts behind or collide with their relationship ids.
        foreach (var old in root.Elements(W + "font").Where(f => (string?)f.Attribute(W + "name") == Family).ToList())
        {
            foreach (var id in old.Elements().Select(e => (string?)e.Attribute(R + "id")).OfType<string>())
                if (table.Parts.Any(p => p.RelationshipId == id)) table.DeletePart(id);
            old.Remove();
        }

        var font = new XElement(W + "font", new XAttribute(W + "name", Family));
        var embedded = new List<string>();
        var missing = new List<string>();
        foreach (var face in Faces)
        {
            if (Find(face.FileName) is not { } path)
            {
                missing.Add(face.FileName);
                continue;
            }
            var key = Guid.NewGuid();
            var part = table.AddNewPart<FontPart>(ObfuscatedFont, null);
            Docx.Write(part, Obfuscate(File.ReadAllBytes(path), key));
            font.Add(new XElement(W + ("embed" + face.Slot),
                new XAttribute(R + "id", table.GetIdOfPart(part)),
                new XAttribute(W + "fontKey", PartTemplate.Braced(key)),
                new XAttribute(W + "subsetted", "false")));
            embedded.Add(face.FileName);
        }
        root.Add(font);
        Docx.Save(table, xml);
        return new FontEmbedResult(Family, FontDirectories, embedded, missing);
    }

    // See FontDirectories. Subdirectory matches are taken in ordinal path order, so the choice
    // does not depend on the file system's enumeration order.
    private string? Find(string fileName)
    {
        foreach (var dir in FontDirectories.Where(Directory.Exists))
        {
            var direct = Path.Combine(dir, fileName);
            if (File.Exists(direct)) return direct;
            if (Path.GetFileName(fileName) != fileName) continue;
            if (Directory.EnumerateFiles(dir, fileName, Subdirectories).Order(StringComparer.Ordinal).FirstOrDefault() is { } nested)
                return nested;
        }
        return null;
    }

    /// <summary>
    /// OOXML font obfuscation (ECMA-376 Part 2, embedded font obfuscation): XOR the first 32 bytes
    /// of the font with the 16 bytes of <paramref name="key"/>, taken from its canonical string form
    /// and applied in reverse order. Symmetric: applying it twice restores the font.
    /// </summary>
    public static byte[] Obfuscate(byte[] font, Guid key)
    {
        var hex = key.ToString("N");
        var k = new byte[16];
        for (var i = 0; i < 16; i++) k[i] = Convert.ToByte(hex.Substring(i * 2, 2), 16);
        var data = (byte[])font.Clone();
        var n = Math.Min(32, data.Length);
        for (var i = 0; i < n; i++) data[i] ^= k[15 - (i % 16)];
        return data;
    }

    private static XDocument NewTable() => new(new XElement(W + "fonts",
        new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName),
        new XAttribute(XNamespace.Xmlns + "r", R.NamespaceName)));

    private static XDocument? TryLoad(OpenXmlPart part)
    {
        try
        {
            return Docx.Load(part) is { Root: { } r } x && r.Name == W + "fonts" ? x : null;
        }
        catch (XmlException)
        {
            return null;
        }
    }
}
